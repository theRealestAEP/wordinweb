import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument, p, W_NS } from "./helpers.js";

// A negative w:top / w:bottom is an ABSOLUTE distance from the page edge, not a
// signed offset. probe-negmargin (wordinweb-parity ca7493d) gives each case its
// own section and its own w:pgMar, and puts a marker with an exact line height
// first in the body so the marker's top IS the body top. It sweeps w:top through
// the negative range against two header heights (the caed-pleading fixture's own
// hRule="exact" 14880tw row, and a 2880tw one) and two w:header distances.
//
// Word's body top came out exactly linear in |w:top| with slope 1 and a constant
// +0.18px residual, and identical to the digit across both header variables:
//
//     w:top (tw)   w:top (px)   Word tall   Word short
//       -2880        -192.00      192.20       192.20
//       -2160        -144.00      144.20       144.20
//       -1440         -96.00       96.18        96.18
//       -1325         -88.33       88.51        88.51
//        -720         -48.00       48.18        48.18
//        -360         -24.00       24.14        24.14
//
// We used to place the body at w:top itself, i.e. ABOVE the top of the page.

const measurer = new ApproxMeasurer();

const PAGE_HEIGHT_TW = 15840;
const PAGE_HEIGHT_PX = PAGE_HEIGHT_TW / 15;

/** Word's measured body top sits a constant +0.18px below |w:top|. */
const TOLERANCE = 0.25;

/** The probe's own header shape: one table row of an exactly known height. */
function headerPart(rowTwips: number): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:hdr ${W_NS}><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>` +
    `<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="${rowTwips}"/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>` +
    `<w:p><w:r><w:t>HDR</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>`
  );
}

/** Lay out one page whose section carries `top`/`bottom` twips, optionally
 * behind a default header whose single row is `headerRowTwips` tall. */
function layoutPage(opts: {
  top: number;
  bottom?: number;
  header?: number;
  headerRowTwips?: number;
}) {
  const headerRef = opts.headerRowTwips
    ? `<w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rId5"/>`
    : "";
  const body =
    p("MARKER") +
    `<w:sectPr>${headerRef}<w:pgSz w:w="12240" w:h="${PAGE_HEIGHT_TW}"/>` +
    `<w:pgMar w:top="${opts.top}" w:right="720" w:bottom="${opts.bottom ?? 1440}" ` +
    `w:left="2088" w:header="${opts.header ?? 432}" w:footer="360" w:gutter="0"/></w:sectPr>`;
  const parts: Record<string, string> = { "word/document.xml": wrapDocument(body) };
  if (opts.headerRowTwips) {
    parts["word/header1.xml"] = headerPart(opts.headerRowTwips);
    parts["word/_rels/document.xml.rels"] =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `</Relationships>`;
  }
  return layoutDocument(DocxDocument.load(makeDocx(parts)), { measurer }).pages[0];
}

describe("negative page margins", () => {
  // The probe's measured w:top rows. Word puts the body top at |w:top| below
  // the top of the page.
  for (const twips of [-2880, -2160, -1440, -1325, -720, -360]) {
    it(`puts the body top at |w:top| for w:top=${twips}`, () => {
      const expected = Math.abs(twips) / 15;
      expect(Math.abs(layoutPage({ top: twips }).bodyTop - expected)).toBeLessThanOrEqual(
        TOLERANCE,
      );
    });
  }

  it("ignores the header height and w:header under a negative top margin", () => {
    // The probe swapped the caed-pleading fixture's own 14880tw exact row for a
    // 2880tw one and moved w:header from 432 to 1440. Word's body top did not
    // move for any of the four combinations.
    const tops = [
      layoutPage({ top: -1325, headerRowTwips: 14880, header: 432 }).bodyTop,
      layoutPage({ top: -1325, headerRowTwips: 2880, header: 432 }).bodyTop,
      layoutPage({ top: -1325, headerRowTwips: 14880, header: 1440 }).bodyTop,
      layoutPage({ top: -1325, headerRowTwips: 2880, header: 1440 }).bodyTop,
    ];
    for (const top of tops) expect(Math.abs(top - 1325 / 15)).toBeLessThanOrEqual(TOLERANCE);
  });

  it("leaves a positive top margin on the header-governed path", () => {
    // Control: above zero the header governs and abs() is a no-op. Headerless,
    // so the known header-height overcharge (#72) stays out of it.
    expect(layoutPage({ top: 1440 }).bodyTop).toBeCloseTo(96, 6);
  });

  it("puts the body bottom at |w:bottom| above the page edge", () => {
    // wild3-template-caed-pleading carries w:bottom="-1267" alongside its
    // negative w:top. The absolute reading is ECMA-consistent and symmetric
    // with the top, and the text is free to run over the footer — but no probe
    // row pins the bottom yet, so this asserts our rule, not a Word measurement.
    const page = layoutPage({ top: -1325, bottom: -1267 });
    expect(page.bodyBottom).toBeCloseTo(PAGE_HEIGHT_PX - 1267 / 15, 6);
  });
});
