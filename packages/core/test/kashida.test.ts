import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { makeDocx, wrapDocument } from "./helpers.js";
import { breakParagraph, type FieldContext } from "../src/layout/inline.js";
import { ApproxMeasurer } from "../src/layout/measure.js";

const FIELDS: FieldContext = {
  pageNumber: () => 1,
  totalPages: () => 1,
  formatPageNumber: (n) => String(n),
};

// Build a bidi (RTL) Arabic paragraph with the given w:jc value. Repeats a
// short Arabic word enough times to wrap across several lines in a ~468pt
// column so justification/kashida behavior is observable.
function arabicPara(jc: string): string {
  const word = "العربية";
  const text = Array(40).fill(word).join(" ");
  return wrapDocument(
    `<w:p><w:pPr><w:bidi/><w:jc w:val="${jc}"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:rtl/></w:rPr>` +
      `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
  );
}

function loadFirstPara(jc: string) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": arabicPara(jc) }));
  const para = doc.sections[0].blocks[0];
  if (para.type !== "paragraph") throw new Error("expected paragraph");
  return { doc, para };
}

describe("kashida / distribute justification", () => {
  it("records the exact justify flavor while keeping alignment=justify", () => {
    for (const jc of ["distribute", "lowKashida", "mediumKashida", "highKashida"] as const) {
      const { para } = loadFirstPara(jc);
      expect(para.props.alignment).toBe("justify");
      expect(para.props.justifyKind).toBe(jc);
    }
  });

  it("plain 'both' justify carries no justifyKind", () => {
    const { para } = loadFirstPara("both");
    expect(para.props.alignment).toBe("justify");
    expect(para.props.justifyKind).toBeUndefined();
  });

  it("distribute stretches the LAST line to the column edge; 'both' leaves it ragged", () => {
    const measurer = new ApproxMeasurer();
    const width = 468 * (4 / 3); // ~468pt column in px
    const lastExtent = (jc: string) => {
      const { doc, para } = loadFirstPara(jc);
      const broken = breakParagraph(doc, measurer, para, width, FIELDS);
      const last = broken.lines[broken.lines.length - 1];
      // Content extent of the last line (independent of which edge it hugs):
      // a justified last line spreads across the whole column, a ragged one
      // only spans its natural word width.
      const l = Math.min(...last.spans.map((s) => s.x));
      const r = Math.max(...last.spans.map((s) => s.x + s.width));
      return r - l;
    };
    const both = lastExtent("both");
    const distribute = lastExtent("distribute");
    // distribute spreads the last line across the column; "both" leaves it ragged.
    expect(distribute).toBeGreaterThan(both + 20);
    expect(distribute).toBeGreaterThan(width - 8);
  });

  it("mediumKashida packs fewer words per line than plain 'both' (glyph elongation)", () => {
    const measurer = new ApproxMeasurer();
    const width = 468 * (4 / 3);
    const lineCount = (jc: string) => {
      const { doc, para } = loadFirstPara(jc);
      return breakParagraph(doc, measurer, para, width, FIELDS).lines.length;
    };
    // The kashida elongation widens packed text, so it needs at least as many
    // lines as plain justify — and medium/high strictly more for this text.
    expect(lineCount("mediumKashida")).toBeGreaterThanOrEqual(lineCount("both"));
    expect(lineCount("highKashida")).toBeGreaterThanOrEqual(lineCount("mediumKashida"));
  });
});
