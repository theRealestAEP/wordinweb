import { describe, expect, it } from "vitest";
import { histogramMatch } from "../src/edit/compare/histogram.js";

const range = (n: number): number[] => Array.from({ length: n }, (_, i) => i);

describe("histogramMatch", () => {
  it("matches everything but the edits, in order", () => {
    const a = ["intro", "x", "body", "y", "end"];
    const b = ["intro", "body", "z", "end"];
    expect(histogramMatch(a, b)).toEqual([
      { a: 0, b: 0 },
      { a: 2, b: 1 },
      { a: 4, b: 3 },
    ]);
  });

  it("aligns a long unchanged run once, not once per row", () => {
    // 200k rows with three edited: growing the run again from each of its
    // rows took 107 s.
    const a = range(200000);
    const b = a.slice();
    for (const k of [50000, 100000, 150000]) b[k] = -k;
    const matches = histogramMatch(a, b);
    expect(matches).toHaveLength(199997);
    expect(matches.every((m) => m.a === m.b)).toBe(true);
  });
});
