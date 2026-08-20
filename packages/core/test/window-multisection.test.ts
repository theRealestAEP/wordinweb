/**
 * Multi-section documents get a page window.
 *
 * The window used to be gated on the incremental relay's envelope, which
 * requires a single section, so one section break cost a 500-page document its
 * whole window: measured 250.9 MB of page model against 11.9 MB windowed.
 *
 * window-determinism.test.ts proves rebuilt pages are byte-identical. This
 * file covers the plumbing either side of that: the window is actually
 * installed, and it is installed on the ASYNC path too, which is the one the
 * editor uses.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, layoutDocumentAsync } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();
const GEOMETRY = `<w:pgSz w:w="7200" w:h="4000"/><w:pgMar w:top="240" w:right="240" w:bottom="240" w:left="240"/>`;

function prose(tag: string, count: number): string {
  return Array.from({ length: count }, (_, i) =>
    p(`${tag}-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet`),
  ).join("");
}

/** `sections` sections of prose, separated by next-page section breaks. */
function doc(sections: number, perSection = 120): DocxDocument {
  const body = Array.from({ length: sections }, (_, s) =>
    prose(`s${s}`, perSection) +
    (s < sections - 1 ? `<w:p><w:pPr><w:sectPr>${GEOMETRY}</w:sectPr></w:pPr></w:p>` : ""),
  ).join("");
  return DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(body + `<w:sectPr>${GEOMETRY}</w:sectPr>`) }),
  );
}

describe("windowing multi-section documents", () => {
  it("installs a window and discards pages behind it", () => {
    const multi = doc(3);
    expect(multi.sections.length).toBe(3);
    const result = layoutDocument(multi, { measurer, windowModel: true });

    expect(result._window).toBeTruthy();
    const discarded = result.pages.filter((page) => page.items.length === 0).length;
    expect(discarded).toBeGreaterThan(result.pages.length / 2);
  });

  it("windows on the async path the editor uses", async () => {
    const result = await layoutDocumentAsync(doc(3), { measurer, windowModel: true });
    expect(result._window).toBeTruthy();
    expect(result.pages.filter((page) => page.items.length === 0).length).toBeGreaterThan(0);
  });

  it("matches a full layout page for page", async () => {
    const full = layoutDocument(doc(3), { measurer });
    const windowed = await layoutDocumentAsync(doc(3), { measurer, windowModel: true });
    expect(windowed.totalPages).toBe(full.totalPages);
    windowed._window!.materialize(windowed.pages.keys());
    const strip = (result: { pages: unknown[] }): string =>
      JSON.stringify(result.pages, (key, value) => (key === "src" || key === "tbl" ? undefined : value));
    expect(strip(windowed)).toBe(strip(full));
  });

  it("still refuses a section that would need column balancing", () => {
    // A rebuild re-enters below layoutSection's balancing pass, so a
    // multi-column section stays outside the window's envelope.
    const columns = `<w:cols w:num="2" w:space="360"/>`;
    const body =
      prose("a", 120) +
      `<w:p><w:pPr><w:sectPr>${GEOMETRY}</w:sectPr></w:pPr></w:p>` +
      prose("b", 120);
    const withColumns = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(body + `<w:sectPr>${columns}${GEOMETRY}</w:sectPr>`),
      }),
    );
    const result = layoutDocument(withColumns, { measurer, windowModel: true });
    expect(result._window).toBeFalsy();
  });
});
