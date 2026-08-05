/**
 * DETERMINISM GATE for the windowed page model.
 *
 * The window keeps positioned items for a handful of pages and rebuilds the
 * rest on demand from capture points. Caret positioning, hit testing and the
 * parity harness all read those items, so a rematerialized page must be
 * byte-identical to the page a full layout produces. This file scrolls a
 * window across each fixture and compares every page it materializes against
 * a full, unwindowed layout of the same document.
 */
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { LayoutResult } from "../src/layout/types.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();

const TALL = `<w:pgSz w:w="7200" w:h="10000"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/>`;
const SHORT = `<w:pgSz w:w="7200" w:h="3000"/><w:pgMar w:top="240" w:right="240" w:bottom="240" w:left="240"/>`;

/** Positioned-item projection of one page. `src`/`tbl` are back-references
 * into the parsed model, not layout output, so they are dropped. */
function pageProjection(result: LayoutResult, pageIndex: number): string {
  return JSON.stringify(result.pages[pageIndex], (key, value) =>
    key === "src" || key === "tbl" ? undefined : value,
  );
}

function docProjection(result: LayoutResult): string {
  return JSON.stringify(result.pages, (key, value) => (key === "src" || key === "tbl" ? undefined : value));
}

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** A table whose rows carry enough text to split across a page boundary. */
function table(tag: string, rows: number, cols: number): string {
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${Math.floor(6000 / cols)}"/>`).join("");
  const body = Array.from({ length: rows }, (_, r) =>
    `<w:tr>` +
    Array.from({ length: cols }, (_, c) =>
      `<w:tc><w:tcPr><w:tcW w:w="${Math.floor(6000 / cols)}" w:type="dxa"/></w:tcPr>` +
      p(`${tag} r${r} c${c} alpha bravo charlie delta echo foxtrot`) +
      `</w:tc>`,
    ).join("") +
    `</w:tr>`,
  ).join("");
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>` +
    `<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>` +
    `<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>` +
    `</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
  );
}

interface Fixture {
  name: string;
  doc: DocxDocument;
  /** Whether this document is inside the windowing envelope at all. */
  windowed: boolean;
}

function fixtures(): Fixture[] {
  const prose = (tag: string, count: number): string =>
    Array.from({ length: count }, (_, i) =>
      p(`${tag}-${i} alpha bravo charlie delta echo foxtrot golf hotel india juliet`),
    ).join("");

  // Long prose with hard page breaks — the plain baseline.
  const paged = Array.from({ length: 40 }, (_, i) =>
    p(`page-${i} alpha bravo charlie delta`) +
    (i < 39 ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` : ""),
  ).join("");

  // Table-heavy: tables large enough to break across pages, interleaved with
  // prose so page tops land both inside and between tables.
  const tables = Array.from({ length: 14 }, (_, i) =>
    table(`t${i}`, 8, 3) + prose(`gap${i}`, 6),
  ).join("");

  // Multi-section: two sections with different page geometry.
  const sectionOne =
    prose("s1", 160) +
    `<w:p><w:pPr><w:sectPr>${TALL}</w:sectPr></w:pPr></w:p>`;
  const multi = sectionOne + prose("s2", 320);

  /** A section-break paragraph closing a section with `props`. */
  const breakPara = (props: string): string => `<w:p><w:pPr><w:sectPr>${props}</w:sectPr></w:pPr></w:p>`;

  // Several sections of alternating geometry. A rebuild has to re-enter part
  // way through and then cross every later boundary, so the resume section
  // index has to be right for points in the middle sections, not just the last.
  const manySections = Array.from({ length: 5 }, (_, i) =>
    prose(`ms${i}`, 70) + breakPara(i % 2 === 0 ? TALL : SHORT),
  ).join("") + prose("ms-last", 70);

  // Continuous breaks: the next section resumes on the SAME page at the current
  // cursor rather than starting a fresh one. That is the boundary path where a
  // rebuild is most likely to disagree with a full run, because the resumed
  // state has to put the cursor back mid-page.
  const continuous = Array.from({ length: 6 }, (_, i) =>
    prose(`cs${i}`, 200) + breakPara(`<w:type w:val="continuous"/>${TALL}`),
  ).join("") + prose("cs-last", 200);

  // Mixed flow: headings, lists, nested prose and tables together. The spaced
  // paragraphs make the gate sensitive to the before/after spacing the resume
  // state has to carry across a capture point.
  const spaced = (tag: string, count: number): string =>
    Array.from({ length: count }, (_, i) =>
      p(
        `${tag}-${i} alpha bravo charlie delta echo foxtrot golf hotel`,
        `<w:pPr><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/></w:pPr>`,
      ),
    ).join("");

  const mixed = Array.from({ length: 16 }, (_, i) =>
    p(`Heading ${i}`, `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`) +
    spaced(`sp${i}`, 3) +
    prose(`m${i}`, 4) +
    table(`mt${i}`, 4, 2) +
    Array.from({ length: 5 }, (_, j) =>
      p(`item ${i}.${j} kilo lima mike november`, `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`),
    ).join(""),
  ).join("");

  return [
    { name: "paged prose", doc: load(paged + `<w:sectPr>${TALL}</w:sectPr>`), windowed: true },
    { name: "dense prose", doc: load(prose("d", 340) + `<w:sectPr>${SHORT}</w:sectPr>`), windowed: true },
    { name: "table-heavy", doc: load(tables + `<w:sectPr>${SHORT}</w:sectPr>`), windowed: true },
    { name: "mixed flow", doc: load(mixed + `<w:sectPr>${SHORT}</w:sectPr>`), windowed: true },
    { name: "multi-section", doc: load(multi + `<w:sectPr>${SHORT}</w:sectPr>`), windowed: true },
    { name: "many sections", doc: load(manySections + `<w:sectPr>${SHORT}</w:sectPr>`), windowed: true },
    { name: "continuous sections", doc: load(continuous + `<w:sectPr>${TALL}</w:sectPr>`), windowed: true },
  ];
}

describe("windowed layout determinism", () => {
  for (const { name, doc, windowed } of fixtures()) {
    describe(name, () => {
      const full = layoutDocument(doc, { measurer });
      const win = layoutDocument(doc, { measurer, windowModel: true });

      it("lays out the same page count as a full layout", () => {
        expect(win.totalPages).toBe(full.totalPages);
        expect(win.totalPages).toBeGreaterThan(20);
      });

      it(windowed ? "installs a page window" : "stays fully materialized outside the windowing envelope", () => {
        expect(Boolean(win._window)).toBe(windowed);
      });

      it("matches a full layout once every page is materialized", () => {
        win._window?.materialize(win.pages.keys());
        expect(docProjection(win)).toBe(docProjection(full));
      });

      if (!windowed) return;

      /** Rebuild `wanted` from capture points and compare it to a full layout.
       * Everything is released first, so no page can pass by still holding the
       * items the original run laid — each one is genuinely rematerialized. */
      const rebuildAndCompare = (wanted: number[], where: string): void => {
        win._window!.releaseExcept([]);
        expect(win._window!.retainedPages().size, "release left pages behind").toBe(0);
        win._window!.materialize(wanted);
        win._window!.releaseExcept(wanted);
        for (const index of wanted) {
          expect(pageProjection(win, index), `page ${index + 1} ${where}`)
            .toBe(pageProjection(full, index));
        }
      };

      it("reproduces every page byte-for-byte under a forward scroll", () => {
        const size = 4;
        for (let top = 0; top < win.totalPages; top += size) {
          const wanted = Array.from({ length: size }, (_, i) => top + i).filter((i) => i < win.totalPages);
          rebuildAndCompare(wanted, `after scrolling to ${top + 1}`);
        }
      });

      it("reproduces every page byte-for-byte under a backward scroll", () => {
        const size = 4;
        for (let top = win.totalPages - size; top >= 0; top -= size) {
          const wanted = Array.from({ length: size }, (_, i) => top + i).filter((i) => i < win.totalPages);
          rebuildAndCompare(wanted, `after scrolling back to ${top + 1}`);
        }
      });

      it("rebuilds a single page in isolation, for every page in the document", () => {
        for (let index = 0; index < win.totalPages; index++) {
          rebuildAndCompare([index], "rebuilt on its own");
        }
      });

      it("reproduces the same pages regardless of the order they are visited", () => {
        // Deterministic pseudo-random jumps: a rebuilt page must depend only on
        // its capture point, never on which pages were materialized before it.
        let seed = 0x2f6e2b1;
        const next = (): number => {
          seed = (seed * 1103515245 + 12345) & 0x7fffffff;
          return seed;
        };
        for (let visit = 0; visit < 24; visit++) {
          const top = next() % win.totalPages;
          const wanted = [top, top + 1, top + 2].filter((i) => i < win.totalPages);
          // Jumps keep whatever the previous window retained, so this also
          // covers materializing into a partially-populated window.
          win._window!.materialize(wanted);
          win._window!.releaseExcept(wanted);
          for (const index of wanted) {
            expect(pageProjection(win, index), `page ${index + 1} after jumping to ${top + 1}`)
              .toBe(pageProjection(full, index));
          }
        }
      });

      it("releases the pages outside the window", () => {
        win._window!.materialize([5, 6, 7]);
        win._window!.releaseExcept([5, 6, 7]);
        expect([...win._window!.retainedPages()].sort((a, b) => a - b)).toEqual([5, 6, 7]);
        const released = win.pages.filter((_, index) => ![5, 6, 7].includes(index));
        expect(released.every((page) => page.items.length === 0)).toBe(true);
      });
    });
  }
});
