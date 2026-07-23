import { describe, expect, it } from "vitest";

/** Doc 14 §3 L2 collab gate: suggesting mode must be REFUSED while the
 * editor emits intents — its mutation paths (revision marks, glyph strikes)
 * do not emit yet, and a non-emitting mutation branch in a live session is
 * the silent-divergence class every fidelity regression in this repo
 * exists to prevent. The audit rule is standing: no mutation branch
 * without an emission, and until suggest intents exist, no suggesting. */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

describe("suggesting mode collab gate", () => {
  it("setSuggesting refuses to enable while onIntent is wired (source-level pin)", () => {
    // A source-level pin (the DocxEditor needs a DOM to instantiate): the
    // gate line must sit at the TOP of setSuggesting, before any state
    // change. If someone removes or reorders it, this fails loudly.
    const src = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), "../src/edit/editor.ts"),
      "utf8",
    );
    const fn = src.slice(src.indexOf("setSuggesting(on: boolean"));
    const gate = fn.indexOf("if (on && this.host.onIntent) return;");
    const firstMutation = fn.indexOf("this.suggesting = on;");
    expect(gate).toBeGreaterThan(-1);
    expect(gate).toBeLessThan(firstMutation);
  });
});
