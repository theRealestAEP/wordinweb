import { describe, expect, it } from "vitest";
import { graphemeBoundaries, graphemeStep } from "../src/edit/grapheme.js";

describe("grapheme boundaries", () => {
  it("treats plain ASCII as one boundary per char", () => {
    expect(graphemeBoundaries("abc")).toEqual([0, 1, 2, 3]);
  });

  it("groups a Devanagari conjunct + matra as one cluster", () => {
    // नमस्ते = न | म | स्ते  (the third cluster is स + virama + त + े)
    const w = "नमस्ते";
    expect(w.length).toBe(6);
    expect(graphemeBoundaries(w)).toEqual([0, 1, 2, 6]);
  });

  it("groups an Arabic base + harakāt as one cluster", () => {
    // بَعْضًا — each consonant carries combining vowel/tanwin marks
    const w = "بَعْضًا";
    const b = graphemeBoundaries(w);
    // fewer clusters than code units (marks fold into their base)
    expect(b.length - 1).toBeLessThan(w.length);
    expect(b[0]).toBe(0);
    expect(b[b.length - 1]).toBe(w.length);
  });

  it("keeps a surrogate pair (astral emoji) intact", () => {
    const w = "a😀b"; // 😀 is a surrogate pair (length 2)
    expect(w.length).toBe(4);
    expect(graphemeBoundaries(w)).toEqual([0, 1, 3, 4]);
  });

  it("empty string yields a single boundary", () => {
    expect(graphemeBoundaries("")).toEqual([0]);
  });
});

describe("grapheme step", () => {
  it("steps forward and back over a Devanagari cluster", () => {
    const w = "नमस्ते"; // boundaries 0,1,2,6
    expect(graphemeStep(w, 2, 1)).toBe(6); // forward jumps the whole conjunct
    expect(graphemeStep(w, 6, -1)).toBe(2); // back returns to the cluster start
  });

  it("returns null at the ends so the caller crosses runs", () => {
    const w = "नमस्ते";
    expect(graphemeStep(w, 6, 1)).toBeNull();
    expect(graphemeStep(w, 0, -1)).toBeNull();
  });

  it("snaps a mid-cluster offset to the nearest boundary in travel direction", () => {
    const w = "नमस्ते"; // cluster स्ते spans 2..6
    expect(graphemeStep(w, 4, 1)).toBe(6); // forward exits the cluster
    expect(graphemeStep(w, 4, -1)).toBe(2); // back enters the cluster start
  });

  it("does not split a surrogate pair", () => {
    const w = "a😀b";
    expect(graphemeStep(w, 1, 1)).toBe(3); // skip both surrogate halves
    expect(graphemeStep(w, 3, -1)).toBe(1);
  });
});
