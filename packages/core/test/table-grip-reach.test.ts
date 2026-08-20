/**
 * #154: a table's grips have to be reachable — the resize grips need a real
 * box, and the move handle needs to accept a press.
 *
 * Both halves were invisible to every existing test because a grip is chrome:
 * it paints nothing into a parity capture (grips are skipped unless the render
 * is interactive), and the editor reaches it through a mousedown on the
 * element, which nothing in a unit test performs.
 *
 * The fixture is BUILT here rather than read from the parity corpus. It used
 * to be an absolute path into a checkout on one machine, which passed there
 * and failed everywhere else — CI has never once run this file. Building it
 * keeps the defect covered on any machine: with `offsetItem`'s `x2 += dx`
 * removed, the nested table below produces four row grips whose far edge sits
 * 45px BEHIND their near edge, which is what these two tests catch.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { PageItem } from "../src/layout/types.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

type Grip = Extract<PageItem, { kind: "grip" }>;

function table(widths: number[], rows: number, body: (row: number, col: number) => string): string {
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const trs = Array.from({ length: rows }, (_, r) => {
    const tcs = widths
      .map((w, c) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${body(r, c)}</w:tc>`)
      .join("");
    return `<w:tr>${tcs}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>`;
}

/**
 * A table inside a cell of another table, so the inner table's items are
 * offset into the outer cell — the only arrangement that moves an item and can
 * therefore leave its far edge behind.
 */
function nestedTableGrips(): Grip[] {
  const inner = table([1200, 1200, 1200], 4, (r, c) => p(`i${r}${c}`));
  // Word requires a paragraph after a nested table inside its cell.
  const outer = table([3000, 3900, 3000], 4, (r, c) =>
    r === 1 && c === 1 ? `${inner}${p("")}` : p(`o${r}${c}`));
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(outer) }));
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
    const all = nestedTableGrips();
    // Named exactly, so a change in what tables offer is a decision rather
    // than a silently weaker test: both tables' rows, both tables' column
    // boundaries, and one move handle.
    const byAxis = all.reduce<Record<string, number>>((acc, grip) => {
      acc[grip.axis] = (acc[grip.axis] ?? 0) + 1;
      return acc;
    }, {});
    expect(byAxis, "the fixture stopped producing grips").toEqual({ move: 1, row: 8, col: 6 });

    const dead = all
      .map((grip) => ({ grip, box: renderedBox(grip) }))
      .filter(({ box }) => box.width <= 0 || box.height <= 0);
    // Before the offsetItem fix: the inner table's row grips came out 45px
    // wide in the negative direction, because a nested table's items were
    // moved by x but their far edge x2 was left in the inner frame's
    // coordinates.
    expect(
      dead.map(({ grip, box }) => `${grip.axis} ${box.width}x${box.height} (x=${grip.x}, x2=${grip.x2})`),
      "grips with no area: nothing can be pressed on them at any z-index",
    ).toEqual([]);
  });

  it("keeps a grip's far edge to the right of its near edge", () => {
    const backwards = nestedTableGrips().filter((grip) => grip.x2 !== undefined && grip.x2 < grip.x);
    // The editor hit-tests the move handle with `lx < x || lx > x2`, which no
    // pointer can satisfy when x2 is behind x — so this is the same defect
    // seen from the drag path rather than from the paint.
    expect(backwards.length, "grips whose x2 is behind their x").toBe(0);
  });
});
