/**
 * #154: a table's grips have to be reachable — the resize grips need a real
 * box, and the move handle needs to accept a press.
 *
 * Both halves were invisible to every existing test because a grip is chrome:
 * it paints nothing into a parity capture (grips are skipped unless the render
 * is interactive), and the editor reaches it through a mousedown on the
 * element, which nothing in a unit test performs.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { PageItem } from "../src/layout/types.js";

/** A nested table, so the inner table's items are offset into the outer cell. */
const NESTED = "/Users/alexpickett/Desktop/Projects/wordinweb-parity/parity/Nested tables.docx";

type Grip = Extract<PageItem, { kind: "grip" }>;

function grips(file: string): Grip[] {
  const doc = DocxDocument.load(new Uint8Array(readFileSync(file)));
  const pages = layoutDocument(doc, new ApproxMeasurer());
  return pages.pages.flatMap((page) => page.items.filter((i): i is Grip => i.kind === "grip"));
}

/** The box dom.ts gives a grip, in the same arithmetic it uses. */
function renderedBox(grip: Grip): { width: number; height: number } {
  if (grip.axis === "col") return { width: 6, height: grip.y2 - grip.y1 };
  if (grip.axis === "row") return { width: (grip.x2 ?? grip.x) - grip.x, height: 6 };
  return { width: 22, height: 22 };
}

describe("#154 · table grips can be pressed", () => {
  it("gives every grip in a nested table a box with area", () => {
    const all = grips(NESTED);
    expect(all.length, "the fixture stopped producing grips").toBeGreaterThan(20);
    const dead = all
      .map((grip) => ({ grip, box: renderedBox(grip) }))
      .filter(({ box }) => box.width <= 0 || box.height <= 0);
    // Before the offsetItem fix: six row grips came out at -74.7 and -191.3px
    // wide, because a nested table's items were moved by x but their far edge
    // x2 was left in the inner frame's coordinates.
    expect(
      dead.map(({ grip, box }) => `${grip.axis} ${box.width}x${box.height} (x=${grip.x}, x2=${grip.x2})`),
      "grips with no area: nothing can be pressed on them at any z-index",
    ).toEqual([]);
  });

  it("keeps a grip's far edge to the right of its near edge", () => {
    const backwards = grips(NESTED).filter((grip) => grip.x2 !== undefined && grip.x2 < grip.x);
    // The editor hit-tests the move handle with `lx < x || lx > x2`, which no
    // pointer can satisfy when x2 is behind x — so this is the same defect
    // seen from the drag path rather than from the paint.
    expect(backwards.length, "grips whose x2 is behind their x").toBe(0);
  });
});
