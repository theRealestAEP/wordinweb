import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";

/**
 * A STRING CHECK ON CSS, deliberately — read why before replacing it.
 *
 * DocxView keeps only the visible pages mounted, and it derives that window
 * from its scroll container's `clientHeight` (`packages/core/src/render/dom.ts`,
 * the `top + Math.max(container.clientHeight, 1)` line). That requires an
 * unbroken height chain from the root: without `height` on html/body/#root, the
 * `height: 100%` inside the app resolves against an auto-height root, the
 * editor's container grows to fit its content rather than scrolling,
 * `clientHeight` becomes the height of the WHOLE document, and nothing is ever
 * culled.
 *
 * The consequence is a 500-page file mounting all 500 pages, with every
 * keystroke relayouting and repainting the lot. That shipped to the deployed
 * demo and was reported as "horrible performance" — the page looks correct,
 * so nothing points at the cause.
 *
 * jsdom cannot catch this: it has no layout engine, so every clientHeight is 0
 * and virtualization is untestable there. A real browser test would catch it
 * but needs Playwright and a large fixture. So this pins the one line whose
 * absence causes it, and fails loudly if someone tidies it away.
 */
const DEMO_HTML = join(__dirname, "../../../examples/anon-share/index.html");

describe("the demo's root height chain (page virtualization depends on it)", () => {
  const css = readFileSync(DEMO_HTML, "utf8");

  it("gives html, body and #root a height, so the editor's container scrolls", () => {
    // Tolerant of ordering and whitespace, strict about all three being present
    // in one rule with a height — any of the three missing breaks the chain.
    const rule = /html\s*,\s*body\s*,\s*#root\s*\{[^}]*height\s*:\s*100%/;
    expect(
      rule.test(css),
      "examples/anon-share/index.html must set `html, body, #root { height: 100% }`. " +
        "Without it DocxView mounts every page of the document instead of the visible ones.",
    ).toBe(true);
  });

  it("does not later reset that height to auto", () => {
    // A `#root { height: auto }` anywhere after the rule above would undo it
    // while leaving the pinned rule present — passing this file's first test
    // and still breaking virtualization.
    expect(/#root\s*\{[^}]*height\s*:\s*auto/.test(css)).toBe(false);
  });
});
