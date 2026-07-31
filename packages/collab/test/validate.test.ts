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
