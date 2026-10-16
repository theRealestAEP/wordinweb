import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, W_NS } from "./helpers.js";

/**
 * An EXACT line taller than its text sits the text at the bottom of the box,
 * with the leading above. Its fit extent runs to the raw font box under that
 * baseline, so a tall exact line whose top fits but whose text would land
 * past the body bottom moves to the next page.
 *
 * probe-sectadvance's shims are such lines (2865tw exact, 11pt text). With a
 * compare's insertions above them, Word's export of the revised document
 * moves BKNP33SHIM from page 3 to page 4 while we painted its text at
 * baseline 1081 of a 1056px page.
 */

const measurer = new ApproxMeasurer();

/** 12240 x 15840 with 1" margins: the body is 96..960 px. */
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr>";

const RPR = '<w:rPr><w:sz w:val="22"/></w:rPr>';

const exact = (twips: number, text: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>${RPR}</w:pPr><w:r>${RPR}<w:t>${text}</w:t></w:r></w:p>`;

/** A filler line leaving `room` px under it, then the 2865tw (191px) shim line. */
function pagesWithRoom(room: number) {
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
        `<w:body>${exact(Math.round((864 - room) * 15), "FILL")}${exact(2865, "SHIM")}${SECT}</w:body></w:document>`,
    }),
  );
  return layoutDocument(doc, { measurer }).pages;
}

function shimOf(pages: ReturnType<typeof pagesWithRoom>) {
  for (let page = 0; page < pages.length; page++) {
    for (const it of pages[page].items) if (it.kind === "text" && it.text === "SHIM") return { page, baseline: it.baseline, top: it.lineTop };
  }
  throw new Error("no SHIM");
}

describe("exact line fit", () => {
  it("keeps a tall exact line whose text fits", () => {
    const shim = shimOf(pagesWithRoom(191 + 4));
    expect(shim.page).toBe(0);
    expect(shim.baseline).toBeLessThan(960);
  });

  it("moves a tall exact line whose top fits but whose text would land past the body bottom", () => {
    // 100px of room: the 191px line's top fits, its bottom-anchored text does not.
    const shim = shimOf(pagesWithRoom(100));
    expect(shim.page).toBe(1);
    expect(shim.top).toBeCloseTo(96, 1);
    expect(shim.baseline).toBeLessThan(960);
  });
});
