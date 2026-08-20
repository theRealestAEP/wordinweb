import { describe, expect, it } from "vitest";
import { validateIntent } from "../src/validate.js";
import { Intent } from "../src/intents.js";

const base = { clientId: "a", clientSeq: 1, base: 0 };

describe("validateIntent", () => {
  it("accepts well-formed intents", () => {
    expect(validateIntent({ ...base, kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "hi" })).toBeNull();
    expect(validateIntent({ ...base, kind: "deleteText", blockId: 1, runId: 2, start: 0, end: 2 })).toBeNull();
  });
  it("rejects an oversized insert", () => {
    expect(validateIntent({ ...base, kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "x".repeat(200_000) })).toMatch(/too long/);
  });
  it("rejects an empty insert and a bad offset", () => {
    expect(validateIntent({ ...base, kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "" })).toMatch(/empty/);
    expect(validateIntent({ ...base, kind: "insertText", at: { blockId: 1, runId: 2, offset: -5 }, text: "x" })).toMatch(/bad offset/);
  });
  it("rejects an empty/negative delete range", () => {
    expect(validateIntent({ ...base, kind: "deleteText", blockId: 1, runId: 2, start: 5, end: 5 })).toMatch(/empty range/);
    expect(validateIntent({ ...base, kind: "deleteText", blockId: 1, runId: 2, start: 5, end: 2 })).toMatch(/empty range/);
  });
  it("rejects a comment that is empty or missing provenance", () => {
    expect(validateIntent({ ...base, kind: "commentRun", runId: 2, text: "", author: "a", date: "d", paraId: "p" } as Intent)).toMatch(/empty/);
    expect(validateIntent({ ...base, kind: "commentRun", runId: 2, text: "x", author: "a" } as unknown as Intent)).toMatch(/provenance/);
  });
});

/**
 * The envelope is what the sequencer orders and de-duplicates by, so it has to
 * be the right SHAPE before any arithmetic on it means anything. A non-number
 * `base` used to pass `base < 0 || base > seq` — both comparisons coerce — and
 * behave like base=head.
 */
describe("intent envelope", () => {
  const body = { kind: "acceptAllRevisions" as const };
  const envelope = { clientId: "c1", clientSeq: 1, base: 0 };

  it("accepts a well-formed envelope", () => {
    expect(validateIntent({ ...body, ...envelope } as Intent)).toBeNull();
  });

  it("refuses a base that is not a non-negative integer", () => {
    for (const base of [{}, "3", -1, 1.5, NaN, null, undefined]) {
      expect(
        validateIntent({ ...body, ...envelope, base } as unknown as Intent),
        `base ${JSON.stringify(base)}`,
      ).toBe("intent: bad base");
    }
  });

  it("refuses a bad clientSeq or clientId", () => {
    expect(validateIntent({ ...body, ...envelope, clientSeq: "1" } as unknown as Intent)).toBe(
      "intent: bad clientSeq",
    );
    expect(validateIntent({ ...body, ...envelope, clientId: "" } as unknown as Intent)).toBe(
      "intent: bad clientId",
    );
    expect(validateIntent({ ...body, ...envelope, clientId: 7 } as unknown as Intent)).toBe(
      "intent: bad clientId",
    );
  });
});
