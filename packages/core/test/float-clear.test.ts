import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx } from "./helpers.js";

/**
 * SQUARE-FLOAT VERTICAL CLEAR — text driven below a wrapSquare float resumes
 * with its line top at EXACTLY the float bottom plus the anchor's distB.
 *
 * Measured from desktop Word via parity's probe-wrapclear: the box height is
 * swept in 3pt steps and Word's first cleared line tracks it 1:1 (no padding,
 * no snapping to the line grid), both when the box leaves no usable strip
 * beside it and when one short word fits beside it and the NEXT line clears.
 * The engine used to add 2px of slack there, which pushed every line and every
 * table rule below the float 1.5pt low for the rest of the page —
 * staging-tblextreme's whole "growing" drift was this one step.
 *
 * The slack was load-bearing: it moved the cleared line past a float-overlap
 * test whose bottom edge was inclusive, so removing it alone would have left
 * the line narrowed by the float it had just cleared. The bottom edge is now
 * exclusive, which is what lets the line sit exactly at the float bottom AND
 * take the full column.
 */

const measurer = new ApproxMeasurer();
const EMU_PER_PT = 12700;
const PX_PER_PT = 4 / 3;

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const BOX_FILL = "#fff2cc";
const TEXT =
  "around the box now Lorem ipsum dolor sit amet consectetur adipiscing elit " +
  "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";

/** Lay out one page: a paragraph carrying an anchored wrapSquare box plus text
 * that has to flow past it. Returns the box's painted rect and the first line
 * that clears it. */
function layoutWithFloat(widthPt: number, heightPt: number, distBPt: number) {
  const cx = Math.round(widthPt * EMU_PER_PT);
  const cy = Math.round(heightPt * EMU_PER_PT);
  const drawing =
    `<w:drawing><wp:anchor distT="0" distB="${Math.round(distBPt * EMU_PER_PT)}" distL="0" distR="0"` +
    ` simplePos="0" relativeHeight="2516502" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="9" name="Shape 9"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>` +
    `</wps:spPr><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>` +
    `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  const body =
    `<w:p><w:r>${drawing}</w:r><w:r><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
    }),
  );
  const page = layoutDocument(doc, { measurer }).pages[0];

  const box = page.items.find(
    (it): it is Extract<typeof it, { kind: "rect" }> =>
      it.kind === "rect" && it.fill?.toLowerCase() === BOX_FILL && it.height > 1,
  );
  if (!box) throw new Error("float box rect not found");
  const boxBottom = box.y + box.height;

  // The first line that starts at or below the float bottom, at the column's
  // left edge — i.e. one that cleared rather than sitting beside the box.
  const cleared = page.items
    .filter(
      (it): it is Extract<typeof it, { kind: "text" }> =>
        it.kind === "text" && it.text.trim().length > 0 && it.lineTop >= boxBottom - 1 && it.x < box.x + 10,
    )
    .sort((a, b) => a.lineTop - b.lineTop)[0];
  if (!cleared) throw new Error("no cleared line found");
  return { boxBottom, clearedTop: cleared.lineTop };
}

describe("a line clearing a square float resumes at exactly float bottom + distB", () => {
  // Wide box: nothing fits beside it, so the clear comes from the
  // no-usable-segment path (probe-wrapclear group A).
  // Narrower box: a ~32pt strip remains and one short word fits beside the
  // box, so the clear comes from the clearY path (group C).
  for (const widthPt of [440, 436]) {
    for (const heightPt of [50, 53, 56, 59]) {
      for (const distBPt of [0, 3.6]) {
        it(`box ${widthPt}x${heightPt}pt, distB ${distBPt}pt`, () => {
          const { boxBottom, clearedTop } = layoutWithFloat(widthPt, heightPt, distBPt);
          expect(clearedTop).toBeCloseTo(boxBottom + distBPt * PX_PER_PT, 5);
        });
      }
    }
  }

  it("tracks the box height 1:1 — no snapping to the line grid", () => {
    const at = (h: number) => layoutWithFloat(440, h, 0).clearedTop;
    // 3pt steps are a fifth of a line: a grid-snapping clear would return the
    // same y for all four heights.
    const tops = [50, 53, 56, 59].map(at);
    for (let i = 1; i < tops.length; i++) {
      expect(tops[i] - tops[i - 1]).toBeCloseTo(3 * PX_PER_PT, 5);
    }
  });
});
