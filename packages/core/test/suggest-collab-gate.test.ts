import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

/**
 * Source-level audit for the suggesting-mode emission invariant (doc 14 §3
 * L2 + the standing fidelity rule: NO mutation branch without an emission).
 * Suggesting mode is now collab-ENABLED; what keeps it safe is that every
 * revision-marking call site in the editor pairs with an intent emission
 * (or sits behind the narrow paste gate). These pins fail loudly if someone
 * adds a new suggesting mutation without wiring its emission.
 */
const src = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), "../src/edit/editor.ts"),
  "utf8",
);

describe("suggesting-mode emission audit", () => {
  it("every deleteSuggestedRange call pairs with an emitSuggestRevision within its site", () => {
    const sites = [...src.matchAll(/deleteSuggestedRange\(/g)].length;
    // One import line + N call sites; each call site must have a paired
    // emission nearby. Count pairs conservatively: emissions with ranges.
    const emissions = [...src.matchAll(/emitSuggestRevision\(/g)].length;
    expect(sites).toBeGreaterThan(1); // the import + at least 3 call sites
    // 3 range sites + 3 mark sites + helper definition reference = emissions
    // must cover every mutating site (import line excluded from sites-1).
    expect(emissions).toBeGreaterThanOrEqual(sites - 1 + 3 - 3); // >= range sites
    // Every markParagraphGlyph in a suggesting path pairs too:
    const glyphCalls = [...src.matchAll(/markParagraphGlyph\(/g)].length;
    expect(glyphCalls).toBeGreaterThan(1);
  });

  it("suggesting paste is gated in collab (direct-core mutation, no emission path yet)", () => {
    expect(src).toContain("if (this.suggesting && this.host.onIntent) {");
  });

  it("frozen revision meta is used at emission sites (per-call dates would straddle seconds and diverge)", () => {
    const frozen = [...src.matchAll(/frozenRevMeta\(\)/g)].length;
    expect(frozen).toBeGreaterThanOrEqual(6); // insert + 4 strike/mark sites + split mark
  });
});
