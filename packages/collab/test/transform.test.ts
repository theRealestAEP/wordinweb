import { describe, expect, it } from "vitest";
import { transformPosition, transformIntent, runEditsOf } from "../src/transform.js";
import { InsertTextIntent, DeleteTextIntent, SplitParagraphIntent, Position } from "../src/intents.js";

const pos = (runId: number, offset: number, blockId = 1): Position => ({ blockId, runId, offset });

function ins(runId: number, offset: number, text: string): InsertTextIntent {
  return { kind: "insertText", clientId: "c", clientSeq: 1, base: 0, at: pos(runId, offset), text };
}
function del(runId: number, start: number, end: number): DeleteTextIntent {
  return { kind: "deleteText", clientId: "c", clientSeq: 1, base: 0, blockId: 1, runId, start, end };
}
function split(runId: number, offset: number, newRunId: number): SplitParagraphIntent {
  return { kind: "splitParagraph", clientId: "c", clientSeq: 1, base: 0, at: pos(runId, offset), newBlockId: 99, newRunId };
}

describe("transformPosition against a prior insert in the same run", () => {
  it("shifts right when the insert is before the position", () => {
    expect(transformPosition(pos(1, 5), [ins(1, 2, "abc")]).offset).toBe(8);
  });
  it("does not shift when the insert is after the position", () => {
    expect(transformPosition(pos(1, 5), [ins(1, 7, "abc")]).offset).toBe(5);
  });
  it("shifts when the insert is exactly at the position (insert lands before caret)", () => {
    // Insert at offset == pos: delEnd (2) <= pos (2) is false only when equal;
    // an insert at the same point pushes the existing caret right.
    expect(transformPosition(pos(1, 2), [ins(1, 2, "XY")]).offset).toBe(4);
  });
  it("ignores inserts in a different run", () => {
    expect(transformPosition(pos(1, 5), [ins(2, 0, "abc")]).offset).toBe(5);
  });
});

describe("transformPosition against a prior delete in the same run", () => {
  it("shifts left when the delete is fully before the position", () => {
    expect(transformPosition(pos(1, 10), [del(1, 2, 5)]).offset).toBe(7); // removed 3
  });
  it("does not shift when the delete is after the position", () => {
    expect(transformPosition(pos(1, 3), [del(1, 5, 9)]).offset).toBe(3);
  });
  it("collapses a position inside the deleted span to the delete point", () => {
    expect(transformPosition(pos(1, 6), [del(1, 4, 9)]).offset).toBe(4);
  });
});

describe("transformPosition against a prior split", () => {
  it("remaps a position after the split point into the new run", () => {
    const p = transformPosition(pos(1, 8), [split(1, 5, 42)]);
    expect(p.runId).toBe(42);
    expect(p.offset).toBe(3); // 8 - 5
  });
  it("leaves a position at or before the split point in the original run", () => {
    expect(transformPosition(pos(1, 5), [split(1, 5, 42)])).toEqual(pos(1, 5));
    expect(transformPosition(pos(1, 2), [split(1, 5, 42)])).toEqual(pos(1, 2));
  });
  it("composes: a remapped position then transforms against a later edit in the new run", () => {
    // Position 8 in run 1 → split at 5 into run 42 → offset 3; then an insert
    // at offset 1 of run 42 shifts it to 5.
    const p = transformPosition(pos(1, 8), [split(1, 5, 42), ins(42, 1, "ab")]);
    expect(p.runId).toBe(42);
    expect(p.offset).toBe(5);
  });
});

describe("transformIntent", () => {
  it("advances base past the intents it was transformed against", () => {
    const i = ins(1, 5, "z");
    const t = transformIntent(i, [ins(1, 0, "aa"), ins(1, 0, "bb")]);
    expect(t.base).toBe(2);
    if (t.kind === "insertText") expect(t.at.offset).toBe(9); // 5 + 2 + 2
  });

  it("transforms both endpoints of a delete", () => {
    const d = del(1, 4, 8);
    const t = transformIntent(d, [ins(1, 0, "xx")]);
    if (t.kind === "deleteText") {
      expect(t.start).toBe(6);
      expect(t.end).toBe(10);
    }
  });

  it("neutralizes a delete whose endpoints split across runs", () => {
    const d = del(1, 3, 9);
    const t = transformIntent(d, [split(1, 6, 50)]); // start stays run 1, end → run 50
    if (t.kind === "deleteText") {
      expect(t.runId).toBe(1);
      expect(t.start).toBe(3);
      expect(t.end).toBe(3); // collapsed → no-op
    }
  });
});

describe("runEditsOf", () => {
  it("summarizes each intent kind", () => {
    expect(runEditsOf(ins(1, 2, "abc"))).toEqual([{ runId: 1, at: 2, del: 0, ins: 3 }]);
    expect(runEditsOf(del(1, 2, 5))).toEqual([{ runId: 1, at: 2, del: 3, ins: 0 }]);
    expect(runEditsOf(split(1, 4, 7))).toEqual([{ runId: 1, at: 4, del: 0, ins: 0, movedToRunId: 7 }]);
  });
});
