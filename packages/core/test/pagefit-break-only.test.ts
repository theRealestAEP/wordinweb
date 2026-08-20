import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { FontMetrics } from "../src/layout/measure.js";
import type { FontSpec } from "../src/layout/types.js";
import { makeDocx, W_NS } from "./helpers.js";

/**
 * The four sweeps of scripts/generate-pagefit-probe.mjs in the parity repo,
 * plus the five-shape sweep of scripts/generate-sectadvance-probe.mjs, as unit
 * cases over our own engine.
 *
 * Each variant fills a page with exact-height paragraphs, adds a shim tuned so
 * an exact amount of room is left, and puts one target paragraph there. The
 * target authors no w:spacing, so its before, line and after all come from
 * w:pPrDefault. Sweeping the room says which of the three the fit test demands.
 *
 * Word and we agree on three of the four: a paragraph needs its space-before
 * and its line, and does not need its space-after. The fourth is the exception
 * this file guards — an EMPTY paragraph whose only run content is a page break
 * demands its SINGLE-SPACED line height, and neither its space-before nor its
 * w:line multiple.
 */

const measurer = new ApproxMeasurer();

/**
 * ApproxMeasurer stands every font's line height in at 1.15em. That is fine
 * for the sweeps below, whose rooms step by 3 px, but the break-only demand IS
 * a line height, and Word bracketed it on real Calibri — 1.2207em, so 16.28 px
 * at 10pt against the approximation's 15.33. Read through the approximation
 * the threshold would pin our arithmetic against a font nobody measured, so
 * the bracket cases use Calibri's own hhea metrics (ascender 1536, descender
 * 512, lineGap 452 over a 2048 em).
 */
class CalibriMeasurer extends ApproxMeasurer {
  metrics(font: FontSpec): FontMetrics {
    return { ascent: font.size * 0.75, descent: font.size * 0.25, lineHeight: font.size * 1.2207 };
  }
}

const calibri = new CalibriMeasurer();

/** No header or footer, so the body runs 96..960 CSS px. */
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr>";

const RPR = '<w:rPr><w:sz w:val="22"/></w:rPr>';

/** phase23-protocol's own defaults: 11pt, before 10pt, after 10pt, line 1.15. */
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W_NS}><w:docDefaults>` +
  `<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
  '<w:pPrDefault><w:pPr><w:spacing w:before="200" w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
  "</w:docDefaults></w:styles>";

/** A paragraph of an exactly known height: no space either side, line exact. */
const exact = (twips: number, text: string, breakBefore = false) =>
  "<w:p><w:pPr>" +
  (breakBefore ? "<w:pageBreakBefore/>" : "") +
  `<w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>${RPR}</w:pPr>` +
  (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : "") +
  "</w:p>";

const BODY_PX = 960 - 96;
const FILLER_TWIPS = 480; // 24 pt = 32 CSS px
const FILLER_PX = 32;
const FILLERS = 25; // 800 px, leaving 64 px for the shim plus the room

const ROOMS = [18, 21, 24, 27, 30, 33, 36, 45];

/** The filled page, the shim, then whatever the sweep puts in the room. */
function variant(room: number, targetXml: string): string {
  let out = "";
  for (let n = 0; n < FILLERS; n++) out += exact(FILLER_TWIPS, "", n === 0);
  out += exact(Math.round((BODY_PX - FILLERS * FILLER_PX - room) * 15), "SHIM");
  return out + targetXml;
}

function pagesOf(bodyXml: string, m: ApproxMeasurer = measurer) {
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
        `<w:body>${bodyXml}${SECT}</w:body></w:document>`,
      "word/styles.xml": STYLES,
    }),
  );
  return layoutDocument(doc, { measurer: m }).pages;
}

/** The 0-based page holding the single text item whose text is `text`. */
function pageOfText(pages: ReturnType<typeof pagesOf>, text: string): number {
  const idx = pages.findIndex((pg) =>
    pg.items.some((i) => i.kind === "text" && i.text === text),
  );
  if (idx < 0) throw new Error(`no text item "${text}"`);
  return idx;
}

/**
 * The smallest swept room at which `fits(room)` holds. The sweeps are
 * monotonic — more room never spills a paragraph that fitted with less — so
 * this single number names the rule. Returns undefined when nothing fits.
 */
function threshold(fits: (room: number) => boolean): number | undefined {
  return ROOMS.find(fits);
}

describe("page fit at the foot of a page", () => {
  // Sweep P: text, no page break. The ordinary rule.
  it("spills a text paragraph until its space-before and line both fit", () => {
    const found = threshold(
      (room) =>
        pageOfText(pagesOf(variant(room, `<w:p><w:pPr>${RPR}</w:pPr><w:r>${RPR}<w:t>TGT</w:t></w:r></w:p>`)), "TGT") === 0,
    );
    expect(found).toBe(33);
  });

  // Sweep Q: the same paragraph plus a page break. A break alone changes nothing.
  it("spills a text paragraph that also carries a page break at the same room", () => {
    const target =
      `<w:p><w:pPr>${RPR}</w:pPr><w:r>${RPR}<w:t>TGT</w:t></w:r>` +
      `<w:r>${RPR}<w:br w:type="page"/></w:r></w:p>`;
    const found = threshold((room) => pageOfText(pagesOf(variant(room, target)), "TGT") === 0);
    expect(found).toBe(33);
  });

  // Sweep S: phase23's block 1395 — EMPTY, only a page break. The exception.
  // Read through a marker: it lands one page after the filled page when the
  // target stays behind, and two pages after when the target spills.
  //
  // Every swept room fits, but "at any room" is NOT what that shows: the sweep
  // floor is 18px and this 11pt paragraph's single-spaced line is 16.9px under
  // ApproxMeasurer (17.9 under real Calibri metrics), so the floor sits just
  // above the threshold the brackets below pin.
  it("keeps an empty paragraph carrying only a page break at every swept room", () => {
    const target =
      `<w:p><w:pPr>${RPR}</w:pPr><w:r>${RPR}<w:br w:type="page"/></w:r></w:p>` +
      exact(FILLER_TWIPS, "AFTER");
    for (const room of ROOMS) {
      expect([room, pageOfText(pagesOf(variant(room, target)), "AFTER")]).toEqual([room, 1]);
    }
  });

  // Sweep T: empty, no page break. Emptiness alone changes nothing — the marker
  // sits at the body top when the empty stayed behind, one paragraph lower when
  // it came along.
  it("spills an empty paragraph with no page break at the ordinary room", () => {
    const target = `<w:p><w:pPr>${RPR}</w:pPr></w:p>` + exact(FILLER_TWIPS, "AFTER");
    const found = threshold((room) => {
      const pages = pagesOf(variant(room, target));
      const marker = pages[1].items.find((i) => i.kind === "text" && i.text === "AFTER");
      if (marker?.kind !== "text") throw new Error("no AFTER marker on page 2");
      return marker.lineTop <= 96 + 0.01; // the empty stayed behind
    });
    expect(found).toBe(33);
  });

  // THE BRACKETS. Shapes BK/B2 of the sectadvance probe and shapes NS/N2 of
  // its no-section control (parity 2ba4f98). Word fits a 10pt break-only
  // paragraph at 17px of room and spills it at 16; at 20pt it fits at 33 and
  // spills at 32. The demand DOUBLES with the font size, so it is a line
  // height and not a constant, and it is the SINGLE-SPACED line: the w:line
  // multiple would make it 18.72px at 10pt (where Word fits at 17) and the
  // space-before would make it 32.05px (which is what the text sweeps above
  // demand and get). The control gives the identical thresholds, so a w:sectPr
  // on the paragraph changes nothing.
  //
  // Read through the marker like sweep S: page 1 when the paragraph stayed
  // behind, page 2 when it spilled onto a page of its own.
  const sectPrInPara =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr>";

  /** The probe target at a given half-point size, with or without the sectPr. */
  const breakOnly = (halfPoints: number, sect: boolean) => {
    const rpr = `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${halfPoints}"/></w:rPr>`;
    return (
      `<w:p><w:pPr>${rpr}${sect ? sectPrInPara : ""}</w:pPr>` +
      `<w:r>${rpr}<w:br w:type="page"/></w:r></w:p>` +
      exact(FILLER_TWIPS, "AFTER")
    );
  };

  const markerPage = (halfPoints: number, sect: boolean, room: number) =>
    pageOfText(pagesOf(variant(room, breakOnly(halfPoints, sect)), calibri), "AFTER");

  for (const sect of [false, true]) {
    const shape = sect ? "with a w:sectPr on it" : "with no w:sectPr";
    it(`charges a 10pt break-only paragraph one 16.3px line ${shape}`, () => {
      expect([markerPage(20, sect, 17), markerPage(20, sect, 16)]).toEqual([1, 2]);
    });

    it(`charges a 20pt break-only paragraph one 32.6px line ${shape}`, () => {
      expect([markerPage(40, sect, 33), markerPage(40, sect, 32)]).toEqual([1, 2]);
    });
  }

  // The coalesce guard. When the sectPr paragraph's line DOES fit, its w:br
  // makes a page and the section then starts ON that page rather than leaving
  // it blank. Suppressing newPage's section-start coalesce for this shape takes
  // BOTH corpus documents carrying it to 24 pages against Word's 23
  // (wild2-legal-ca-agreement p77, wild2-med-nccih-protocol p17), so today's
  // coalesce is right for every case we can lay out.
  //
  // Read at room 45 deliberately. At an insufficient room the paragraph spills
  // and the page it lands on is blank because the mark is invisible, not
  // because of anything the coalesce did - that page says nothing about this
  // rule.
  it("starts a section on the page its own paragraph's break made", () => {
    const target =
      `<w:p><w:pPr>${RPR}${sectPrInPara}</w:pPr><w:r>${RPR}<w:br w:type="page"/></w:r></w:p>` +
      exact(FILLER_TWIPS, "AFTER");
    const pages = pagesOf(variant(45, target));
    const blank = pages.filter((pg) => pg.items.filter((i) => i.kind === "text").length === 0);
    expect([pages.length, blank.length]).toEqual([2, 0]);
  });
});
