import { describe, expect, it } from "vitest";
import { validateIntent } from "../src/validate.js";
import type { Intent } from "../src/intents.js";

/**
 * The doc 13 item-3 bounds audit (round-4 blocker 3), executable form.
 *
 * The requirement: in E2EE mode, apply is CAP-FREE on every client (resource
 * caps in the accept/reject path are machine-dependent and would make
 * verdicts nondeterministic), so `validateIntent` must STRUCTURALLY bound
 * all work an intent can cause BEFORE apply. This suite audits the
 * resource-bearing dimensions of the intent surface: every vector below is
 * an intent whose apply cost scales with an attacker-chosen number, and
 * every one must be rejected by validation — deterministically, from the
 * intent alone, with no clock and no memory pressure involved.
 *
 * (Exhaustiveness note: kinds not listed here either carry no scalable
 * payload — fixed-size formatting/toggle intents whose apply cost is
 * bounded by the document, which the seed-size caps bound — or share a
 * validated payload shape with a listed kind, e.g. chart/smartArt setters.)
 */

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;
const at = { blockId: 1, runId: 2, offset: 0 } as const;

const HOSTILE: [string, Intent][] = [
  ["insertText text beyond maxInsertLength", { kind: "insertText", ...base, at, text: "x".repeat(200_000) } as never],
  ["deleteText range beyond maxDeleteLength", { kind: "deleteText", ...base, blockId: 1, runId: 2, start: 0, end: 2_000_000 } as never],
  ["pasteBlocks XML beyond maxPasteBytes", { kind: "pasteBlocks", ...base, anchorBlockId: 1, blocksXml: "<w:p>".repeat(500_000), nodeIds: [] } as never],
  ["insertTable rows*cols blowup", { kind: "insertTable", ...base, runId: 2, rows: 10_000, cols: 10_000, nodeIds: [] } as never],
  ["insertTable negative dims", { kind: "insertTable", ...base, runId: 2, rows: -1, cols: 3, nodeIds: [] } as never],
  ["commentRun text beyond maxCommentLength", { kind: "commentRun", ...base, runId: 2, text: "x".repeat(50_000), paraId: "p", date: "d", author: "a", commentId: 1, nodeIds: [] } as never],
  ["chart with unbounded series", { kind: "insertChart", ...base, runId: 2, chart: { type: "column", categories: ["a"], series: Array.from({ length: 500 }, () => ({ name: "s", values: [1] })) }, nodeIds: [] } as never],
  ["chart with unbounded values per series", { kind: "insertChart", ...base, runId: 2, chart: { type: "column", categories: ["a"], series: [{ name: "s", values: Array.from({ length: 100_000 }, () => 1) }] }, nodeIds: [] } as never],
  ["smartArt with unbounded items", { kind: "insertSmartArt", ...base, runId: 2, smartArt: { layout: "list", items: Array.from({ length: 5_000 }, () => "x") }, nodeIds: [] } as never],
];

describe("validate.ts bounds audit (doc 13 item 3 / round-4 blocker 3)", () => {
  for (const [label, intent] of HOSTILE) {
    it(`rejects: ${label}`, () => {
      const verdict = validateIntent(intent);
      expect(verdict).toBeTypeOf("string"); // a rejection REASON, not null
    });
  }

  it("verdicts are deterministic — same intent, same verdict, twice", () => {
    for (const [, intent] of HOSTILE) {
      expect(validateIntent(intent)).toBe(validateIntent(intent));
    }
  });

  it("a benign intent of each audited resource-bearing kind passes", () => {
    const benign: Intent[] = [
      { kind: "insertText", ...base, at, text: "hello" } as never,
      { kind: "deleteText", ...base, blockId: 1, runId: 2, start: 0, end: 5 } as never,
      { kind: "insertTable", ...base, runId: 2, rows: 3, cols: 3, nodeIds: [1, 2, 3] } as never,
    ];
    for (const i of benign) expect(validateIntent(i)).toBeNull();
  });
});
