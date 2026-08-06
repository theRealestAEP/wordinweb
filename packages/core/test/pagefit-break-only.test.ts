import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, W_NS } from "./helpers.js";

/**
 * The four sweeps of scripts/generate-pagefit-probe.mjs in the parity repo,
 * as unit cases over our own engine.
 *
 * Each variant fills a page with exact-height paragraphs, adds a shim tuned so
 * an exact amount of room is left, and puts one target paragraph there. The
 * target authors no w:spacing, so its before, line and after all come from
 * w:pPrDefault. Sweeping the room says which of the three the fit test demands.
 *
 * Word and we agree on three of the four: a paragraph needs its space-before
 * and its line, and does not need its space-after. The fourth is the exception
 * this file guards — an EMPTY paragraph whose only run content is a page break
 * stays on the current page at any room at all.
 */

const measurer = new ApproxMeasurer();

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

function pagesOf(bodyXml: string) {
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
        `<w:body>${bodyXml}${SECT}</w:body></w:document>`,
      "word/styles.xml": STYLES,
    }),
  );
  return layoutDocument(doc, { measurer }).pages;
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
  it("keeps an empty paragraph carrying only a page break at any room", () => {
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
});
