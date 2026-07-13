import { describe, expect, it } from "vitest";
import { resolveField, type FieldContext } from "../src/layout/inline.js";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument, W_NS } from "./helpers.js";

const measurer = new ApproxMeasurer();

/** A minimal FieldContext; individual tests override the members they exercise. */
function ctx(over: Partial<FieldContext> = {}): FieldContext {
  return {
    pageNumber: () => 1,
    totalPages: () => 2,
    formatPageNumber: (n) => String(n),
    ...over,
  };
}

/** A legacy complex field: begin / instrText / separate / cached result / end. */
function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function bodyText(body: string): string {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
  const result = layoutDocument(doc, { measurer });
  return result.pages
    .flatMap((pg) => pg.items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : "")))
    .join("");
}

describe("PAGE/NUMPAGES general-formatting (\\*) switch", () => {
  it("PAGE \\* roman overrides the section decimal format", () => {
    // pageNumber 1, section formats decimal — the \* switch forces roman.
    expect(resolveField("PAGE \\* roman", "9", ctx({ pageNumber: () => 1 }))).toBe("i");
  });
  it("PAGE \\* ArabicDash wraps the number in hyphens", () => {
    expect(resolveField("PAGE \\* ArabicDash", "x", ctx({ pageNumber: () => 1 }))).toBe("- 1 -");
  });
  it("plain PAGE keeps the section format", () => {
    expect(resolveField("PAGE", "x", ctx({ pageNumber: () => 3, formatPageNumber: (n) => `p${n}` }))).toBe("p3");
  });
  it("NUMPAGES \\* roman formats the total", () => {
    expect(resolveField("NUMPAGES \\* roman", "9", ctx({ totalPages: () => 2 }))).toBe("ii");
  });
  it("NUMPAGES with a non-numeric switch keeps the plain total", () => {
    expect(resolveField("NUMPAGES \\* MERGEFORMAT", "9", ctx({ totalPages: () => 4 }))).toBe("4");
  });
});

describe("REF recompute", () => {
  it("REF \\h to an empty (captured) bookmark shows nothing, not the stale cache", () => {
    expect(resolveField("REF bk \\h", "Table 1", ctx({ refText: () => "" }))).toBe("");
  });
  it("REF \\h to a known bookmark re-renders its text over the cache", () => {
    expect(resolveField("REF bk \\h", "STALE", ctx({ refText: () => "Introduction" }))).toBe("Introduction");
  });
  it("REF to an uncaptured bookmark keeps its cache", () => {
    expect(resolveField("REF bk \\h", "cached", ctx({ refText: () => undefined }))).toBe("cached");
  });
  it("REF \\p paints the relative position, not the bookmark text", () => {
    const key = {};
    expect(
      resolveField("REF bk \\h \\p", "Table 1 below", ctx({ refPosition: () => "above" }), key),
    ).toBe("above");
  });
  it("REF \\r paints the paragraph number", () => {
    const key = {};
    expect(resolveField("REF bk \\r", "1", ctx({ refParaNumber: () => "0" }), key)).toBe("0");
  });
  it("a \\p ref with no recorded position keeps its cache", () => {
    expect(resolveField("REF bk \\p", "below", ctx({ refPosition: () => undefined }), {})).toBe("below");
  });
});

describe("STYLEREF recompute", () => {
  it("plain STYLEREF resolves the first-on-page style text over the cache", () => {
    expect(
      resolveField('STYLEREF "Heading 1" \\* MERGEFORMAT', "Chapter One", ctx({ styleRef: (n, last) => (n === "Heading 1" && !last ? "Chapter Four" : undefined) })),
    ).toBe("Chapter Four");
  });
  it("STYLEREF \\l asks for the last-on-page paragraph", () => {
    expect(
      resolveField('STYLEREF "Heading 1" \\l \\* MERGEFORMAT', "Chapter One", ctx({ styleRef: (_n, last) => (last ? "Chapter Five" : "Chapter Four") })),
    ).toBe("Chapter Five");
  });
  it("STYLEREF keeps the cache when the resolver returns undefined (no match)", () => {
    expect(resolveField('STYLEREF "Heading 1"', "cached", ctx({ styleRef: () => undefined }))).toBe("cached");
  });
  it("STYLEREF with no resolver (body field) keeps the cache", () => {
    expect(resolveField('STYLEREF "Heading 1"', "cached", ctx())).toBe("cached");
  });
});

describe("index-xrefs style REF fields end-to-end", () => {
  // A zero-length bookmark (start immediately followed by end) sitting before
  // an unnumbered caption, referenced later by \h / \r / \p. Word recomputes
  // all three: empty text, paragraph number 0, position "above".
  const body =
    `<w:p><w:bookmarkStart w:id="1" w:name="bk1"/><w:bookmarkEnd w:id="1"/>` +
    `<w:r><w:t xml:space="preserve">Caption</w:t></w:r></w:p>` +
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">[</w:t></w:r>` +
    field("REF bk1 \\h", "STALEREF") +
    `<w:r><w:t xml:space="preserve">|</w:t></w:r>` +
    field("REF bk1 \\r", "7") +
    `<w:r><w:t xml:space="preserve">|</w:t></w:r>` +
    field("REF bk1 \\p", "below") +
    `<w:r><w:t xml:space="preserve">]</w:t></w:r>` +
    `</w:p>`;

  it("empty \\h => '', \\r => '0', \\p => 'above'", () => {
    expect(bodyText(body)).toBe("Caption[|0|above]");
  });
});

describe("checkbox ballot glyph routing", () => {
  it("routes ballot glyphs to MS Gothic via paintFamily (Latin metrics kept)", () => {
    const body =
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>` +
      `<w:t xml:space="preserve">A☒B</w:t></w:r></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const result = layoutDocument(doc, { measurer });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const ballot = items.find((i) => i.kind === "text" && i.text.includes("☒"));
    const latin = items.find((i) => i.kind === "text" && i.text === "A");
    expect(ballot && ballot.kind === "text" ? ballot.font.paintFamily : undefined).toBe("MS Gothic");
    // The Latin neighbour is NOT rerouted.
    expect(latin && latin.kind === "text" ? latin.font.paintFamily : undefined).toBeUndefined();
    // Metrics stay on the run's own family so the line box is not inflated.
    expect(ballot && ballot.kind === "text" ? ballot.font.family : "").toMatch(/calibri/i);
  });
});

describe("DATE/TIME re-evaluation (Word recomputes on open; cache is stale)", () => {
  it("DATE \\@ ISO picture renders the CURRENT date, not the cache", () => {
    const now = new Date();
    const iso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
    expect(resolveField('DATE \\@ "yyyy-MM-dd"', "1999-01-01", ctx())).toBe(iso);
  });
  it("TIME am/pm token renders the designator UPPERCASE like Word", () => {
    expect(resolveField('TIME \\@ "h:mm am/pm"', "3:07 pm", ctx())).toMatch(/^\d{1,2}:\d{2} [AP]M$/);
  });
  it("month (M) and minute (m) stay case-distinguished in mixed pictures", () => {
    const out = resolveField('DATE \\@ "M/d/yy H:mm"', "x", ctx());
    const now = new Date();
    expect(out.startsWith(`${now.getMonth() + 1}/${now.getDate()}/`)).toBe(true);
    expect(out).toMatch(/ \d{1,2}:\d{2}$/);
  });
  it("CREATEDATE keeps its stored cache (a document moment, not the clock)", () => {
    expect(resolveField("CREATEDATE", "May 5, 2020", ctx())).toBe("May 5, 2020");
  });
});
