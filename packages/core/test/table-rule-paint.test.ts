// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** Sum of the translateY() terms in a transform list, in px. */
function translateY(transform: string): number {
  let total = 0;
  for (const m of transform.matchAll(/translateY\((-?[\d.]+)px\)/g)) total += Number(m[1]);
  return total;
}

describe("table rule paint placement", () => {
  // Browsers round a painted box's position to a whole CSS pixel, so a rule
  // whose layout position is fractional lands up to half a pixel off — two
  // device pixels at the parity renderer's 2x scale — and neighbouring rules
  // in one table round opposite ways, wobbling Word's even grid. Transforms
  // are not rounded, so every rule must sit at an INTEGER `top` with the
  // remainder carried in translateY.
  it("keeps a rule's fractional position in a transform, not in top", () => {
    const row = (text: string) =>
      `<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6000"/></w:tcPr>${p(text)}</w:tc></w:tr>`;
    // spacing line=253 (12.65pt) makes the row pitch fractional in CSS px, so
    // the boundaries land off the whole-pixel grid the way Word's do.
    const lead =
      `<w:p><w:pPr><w:spacing w:line="253" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>lead</w:t></w:r></w:p>`;
    const table = `<w:tbl>
      <w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:color="000000"/>
      </w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      ${row("one")}${row("two")}${row("three")}
    </w:tbl>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(lead + table) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const container = document.createElement("div");
    document.body.appendChild(container);
    renderToDom(doc, result, container);

    const rules = [...container.querySelectorAll<HTMLElement>("[data-dxw-edge]")].filter(
      (el) => Number(parseFloat(el.style.width)) > 3,
    );
    expect(rules.length).toBeGreaterThanOrEqual(4);

    const layoutRules = result.pages[0].items
      .filter((item) => item.kind === "edge" && item.y1 === item.y2)
      .map((item) => (item.kind === "edge" ? item.y1 : 0))
      .sort((a, b) => a - b);
    // The boundaries this test relies on really are off the pixel grid — a
    // whole-pixel grid would make the assertion below vacuous.
    expect(layoutRules.some((y) => !Number.isInteger(y))).toBe(true);

    const placed: number[] = [];
    for (const el of rules) {
      const top = parseFloat(el.style.top);
      // The browser rounds `top`; only a whole number survives paint intact.
      expect(Number.isInteger(top)).toBe(true);
      placed.push(top + translateY(el.style.transform));
    }
    placed.sort((a, b) => a - b);

    // Every rule still paints at its layout position: half the rule's painted
    // width above the boundary, to the last fraction.
    const half = parseFloat(rules[0].style.height) / 2;
    for (const [i, y] of placed.entries()) {
      expect(y + half).toBeCloseTo(layoutRules[i], 6);
    }
    // ...and the pitch Word laid down is preserved exactly, with no rule
    // pulled a pixel toward its neighbour.
    for (let i = 1; i < placed.length; i++) {
      expect(placed[i] - placed[i - 1]).toBeCloseTo(layoutRules[i] - layoutRules[i - 1], 6);
    }
  });
});
