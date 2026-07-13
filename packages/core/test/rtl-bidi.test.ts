import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { makeDocx, wrapDocument } from "./helpers.js";
import { breakParagraph, type FieldContext, type LineSpan } from "../src/layout/inline.js";
import { ApproxMeasurer } from "../src/layout/measure.js";

const FIELDS: FieldContext = {
  pageNumber: () => 1,
  totalPages: () => 1,
  formatPageNumber: (n) => String(n),
};

const WIDTH = 468 * (4 / 3); // ~468pt (9360tw) content column in px

function loadPara(bodyXml: string, extraParts: Record<string, string> = {}) {
  const doc = DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(bodyXml), ...extraParts }),
  );
  const para = doc.sections[0].blocks[0];
  if (para.type !== "paragraph") throw new Error("expected paragraph");
  return { doc, para };
}

function breakBody(bodyXml: string, extraParts?: Record<string, string>) {
  const { doc, para } = loadPara(bodyXml, extraParts);
  const measurer = new ApproxMeasurer();
  return breakParagraph(doc, measurer, para, WIDTH, FIELDS);
}

// Left edge of the first span whose text equals `t` (spans are laid in visual
// order, so x is the on-page position).
function spanX(spans: LineSpan[], t: string): number {
  const s = spans.find((sp) => sp.text === t);
  if (!s) throw new Error(`no span "${t}" in [${spans.map((sp) => JSON.stringify(sp.text)).join(", ")}]`);
  return s.x;
}
function findSpan(spans: LineSpan[], t: string): LineSpan {
  const s = spans.find((sp) => sp.text === t);
  if (!s) throw new Error(`no span "${t}"`);
  return s;
}

describe("RTL bidi embedding (UAX#9) inside a w:rtl run", () => {
  // A single w:rtl run mixing Arabic, Latin words, and European numbers, like
  // probe2-arabic-rtl's "99.9" line. Latin/number islands must be laid LTR and
  // placed in RTL visual order; Arabic punctuation stays on the RTL side.
  const body =
    `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr>` +
    `<w:t xml:space="preserve">الإصدار v2.0 يدعم Unicode 15 و ISO 8601 بنسبة 99.9٪ من الحالات.</w:t>` +
    `</w:r></w:p>`;

  it("lays Latin word + European number as one LTR island (Unicode before 15, ISO before 8601)", () => {
    const { lines } = breakBody(body);
    const spans = lines[0].spans;
    // Within an LTR island the reading order runs left-to-right.
    expect(spanX(spans, "Unicode")).toBeLessThan(spanX(spans, "15"));
    expect(spanX(spans, "ISO")).toBeLessThan(spanX(spans, "8601"));
    // v2.0 (letter + decimal) is a single LTR run.
    expect(findSpan(spans, "v2.0").rtl).toBeFalsy();
  });

  it("flags Latin/number islands LTR and Arabic RTL", () => {
    const spans = breakBody(body).lines[0].spans;
    for (const ltr of ["Unicode", "15", "ISO", "8601", "v2.0", "99.9"]) {
      expect(findSpan(spans, ltr).rtl).toBeFalsy();
    }
    for (const rtl of ["يدعم", "الحالات."]) {
      expect(findSpan(spans, rtl).rtl).toBeTruthy();
    }
  });

  it("keeps the Arabic percent sign on the RTL side of its number (٪ left of 99.9)", () => {
    // U+066A is bidi-class ET but Word treats it as Arabic: it lands after the
    // number in reading order, i.e. to the LEFT of the LTR "99.9" island.
    const spans = breakBody(body).lines[0].spans;
    expect(spanX(spans, "٪")).toBeLessThan(spanX(spans, "99.9"));
    expect(findSpan(spans, "٪").rtl).toBeTruthy();
  });

  it("places the logically-first Arabic word at the visual right (RTL order)", () => {
    const spans = breakBody(body).lines[0].spans;
    // "الإصدار" is the first word read; in RTL it is the rightmost on the line.
    const maxX = Math.max(...spans.map((s) => s.x + s.width));
    expect(findSpan(spans, "الإصدار").x + findSpan(spans, "الإصدار").width).toBeCloseTo(maxX, 0);
  });
});

describe("RTL left tab near the right edge (probe2-arabic-rtl tab line)", () => {
  // bidi + jc=right, an explicit LEFT tab stop at 9000tw (~600px) — close to the
  // ~624px content edge. Word leaves the pre-tab text on its own flush-left line
  // and wraps the post-tab text into the narrow trailing column.
  const body =
    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="9000"/></w:tabs><w:bidi/><w:jc w:val="right"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr><w:t xml:space="preserve">البند الأول</w:t></w:r>` +
    `<w:r><w:tab/></w:r>` +
    `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr><w:t xml:space="preserve">صفحة ١</w:t></w:r>` +
    `</w:p>`;

  it("wraps to multiple flush-left lines instead of one overflowing line", () => {
    const { lines } = breakBody(body);
    expect(lines.length).toBeGreaterThanOrEqual(3);
    // Every visual line starts flush-left (near x=0), never pushed right by a
    // full-width tab gap.
    for (const line of lines) {
      const minX = Math.min(...line.spans.filter((s) => s.text.trim()).map((s) => s.x));
      expect(minX).toBeLessThan(6);
    }
  });

  it("keeps the pre-tab segment alone on the first line", () => {
    const { lines } = breakBody(body);
    const first = lines[0].spans.filter((s) => s.text.trim() && s.text !== "\t");
    const joined = first.map((s) => s.text).join("");
    // "البند الأول" (its two words, in visual RTL order) and nothing from "صفحة".
    expect(joined).not.toContain("صفح");
    expect(joined).toContain("البند");
  });
});

describe("RTL abjad list marker gap", () => {
  // A right-aligned abjad marker ("أ-") with a hanging indent (ind left=720,
  // hanging=360) and a tab suffix — Word keeps a full-hanging gap between the
  // marker and the list text, which reorderVisual must preserve.
  const body =
    `<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/><w:bidi/><w:jc w:val="right"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr>` +
    `<w:t xml:space="preserve">العنصر الأول</w:t></w:r></w:p>`;
  const label = {
    text: "أ-",
    props: { rtl: true, fontComplex: "Arial" },
    suffix: "tab",
    alignment: "right",
  } as unknown as Parameters<typeof breakParagraph>[5];

  it("emits a gap between the abjad marker and the text after reordering", () => {
    const { doc, para } = loadPara(body);
    const spans = breakParagraph(doc, new ApproxMeasurer(), para, WIDTH, FIELDS, label).lines[0].spans;
    const marker = findSpan(spans, "أ-");
    // The marker sits to the RIGHT of the (reordered) body text with a real gap:
    // some content ends left of the marker's start.
    const bodyRight = Math.max(
      ...spans.filter((s) => s !== marker && s.text.trim() && s.text !== "\t").map((s) => s.x + s.width),
    );
    expect(marker.x).toBeGreaterThan(bodyRight + 4); // > ~4px suffix-tab gap
    expect(marker.rtl).toBeTruthy();
  });
});
