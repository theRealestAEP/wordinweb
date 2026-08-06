import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, W_NS } from "./helpers.js";

/**
 * A w:docGrid(lines) TEXT line whose font line exceeds the grid pitch snaps up
 * to a whole number of pitches, and its glyph box sits CENTERED in that band.
 * The half-leading BELOW the glyph box is ordinary leading: Word lets it hang
 * past the bottom margin, exactly as it does for the leading of an ungridded
 * line. Only the glyph box has to fit.
 *
 * Measured on wild2-math-eq-as-images page 7 (parity abaaad1, its Word PDF read
 * by scripts/pdf-page-geometry.py). The page's final paragraph is two SimSun
 * 14pt lines on a 15.6pt grid, so each line snaps to 2 pitches = 41.60 px:
 *
 *     Word baselines   976.63 and 1018.30   (pitch 41.67, ours to 0.1 px)
 *     glyph box bottom 1022.33
 *     body bottom      1026.53
 *
 * Word keeps both lines on the page with 4.20 px to spare, although the second
 * line's 41.60 px band runs 15.80 px past the bottom. We charged the whole band
 * as the fit extent, refused the line, and — two lines under widowControl being
 * unsplittable — moved the whole paragraph to a page Word does not have.
 *
 * The same overcharge is what engine.ts's planBreaks comment recorded as an
 * effective bottom sitting ~14 px high: the deficit was never in the bottom
 * (updateBottom, the note reserves and the banner all read their nominal values
 * on these pages) but in the demand.
 */

const measurer = new ApproxMeasurer();

/** 12240 x 15840 with 1" margins, on a 312tw (15.6pt = 20.8px) line grid. */
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:docGrid w:type="lines" w:linePitch="312"/>' +
  "</w:sectPr>";

/** The snap is a legacy-mode behavior, so the fixture's own compat 12. */
const SETTINGS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>` +
  '<w:compat><w:compatSetting w:name="compatibilityMode" ' +
  'w:uri="http://schemas.microsoft.com/office/word" w:val="12"/></w:compat></w:settings>';

const PITCH_PX = 312 / 15;

/** 14pt: a 21.47px font line against the 20.8px pitch, so every line snaps. */
const RPR = '<w:rPr><w:sz w:val="28"/></w:rPr>';

/** A paragraph of an exactly known height: no space either side, line exact.
 * An exact line never snaps to the grid, so the filler stack is unaffected. */
const exact = (twips: number, text: string) =>
  "<w:p><w:pPr>" +
  `<w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>${RPR}</w:pPr>` +
  (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : "") +
  "</w:p>";

/** Two snapped lines under widow control, split by an explicit break. */
const TARGET =
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>${RPR}</w:pPr>` +
  `<w:r>${RPR}<w:t>ALPHA</w:t></w:r><w:r>${RPR}<w:br/></w:r><w:r>${RPR}<w:t>BETA</w:t></w:r></w:p>`;

function pagesWithRoom(room: number) {
  // The section's first page reserves four grid rows above the body.
  const bodyTop = 96 + 4 * PITCH_PX;
  const shim = Math.round((960 - bodyTop - room) * 15);
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
        `<w:body>${exact(shim, "SHIM")}${TARGET}${SECT}</w:body></w:document>`,
      "word/settings.xml": SETTINGS,
    }),
  );
  return layoutDocument(doc, { measurer }).pages;
}

/** The 0-based page holding the text item whose text is `text`. */
function pageOfText(pages: ReturnType<typeof pagesWithRoom>, text: string): number {
  const idx = pages.findIndex((pg) => pg.items.some((i) => i.kind === "text" && i.text === text));
  if (idx < 0) throw new Error(`no text item "${text}"`);
  return idx;
}

/** Line top, height and painted baseline of the line carrying `text`. */
function lineOf(pages: ReturnType<typeof pagesWithRoom>, text: string) {
  for (const pg of pages) {
    for (const it of pg.items) {
      if (it.kind === "text" && it.text === text) {
        return { top: it.lineTop, height: it.lineHeight, baseline: it.baseline };
      }
    }
  }
  throw new Error(`no text item "${text}"`);
}

describe("docGrid text-snap line fit", () => {
  it("snaps a 14pt line to two grid pitches", () => {
    const line = lineOf(pagesWithRoom(400), "ALPHA");
    expect(line.height).toBeCloseTo(2 * PITCH_PX, 2);
    // Glyph box centered in the band: the baseline sits its ascent below the
    // top half-leading (26.87 px raw, painted on the quarter-point grid), and
    // the same half-leading is left underneath.
    expect(line.baseline - line.top).toBeCloseTo(27.0, 2);
  });

  it("keeps both lines when only the glyph box fits", () => {
    // Room for the first line's whole band plus the second line's glyph box,
    // with the 4.20 px Word had on eq-as-images p7 still to spare. The second
    // band itself runs 5.87 px past the bottom.
    const pages = pagesWithRoom(2 * PITCH_PX + 31.53 + 4.2);
    expect(pageOfText(pages, "ALPHA")).toBe(0);
    expect(pageOfText(pages, "BETA")).toBe(0);
    const second = lineOf(pages, "BETA");
    expect(second.top + second.height).toBeGreaterThan(960);
  });

  it("still spills when the glyph box itself crosses the bottom", () => {
    // One px less room than the glyph box needs. Two lines under widowControl
    // are unsplittable, so the whole paragraph moves.
    const pages = pagesWithRoom(2 * PITCH_PX + 31.53 - 1);
    expect(pageOfText(pages, "ALPHA")).toBe(1);
    expect(pageOfText(pages, "BETA")).toBe(1);
  });
});
