import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

describe("table boundary geometry", () => {
  it("places cell content and surrounding flow inside the outer border halves", () => {
    const table = `<w:tbl>
      <w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="000000"/>
        <w:bottom w:val="single" w:sz="6" w:color="000000"/>
        <w:left w:val="single" w:sz="6" w:color="000000"/>
        <w:right w:val="single" w:sz="6" w:color="000000"/>
      </w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6000"/></w:tcPr>${p("cell")}</w:tc></w:tr>
    </w:tbl>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(p("lead") + table + p("after")) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const items = result.pages[0].items;
    const text = (value: string) => {
      const item = items.find((candidate) => candidate.kind === "text" && candidate.text === value);
      if (item?.kind !== "text") throw new Error(`missing ${value}`);
      return item;
    };
    const rules = items
      .filter((item) => item.kind === "edge" && item.y1 === item.y2)
      .sort((a, b) => (a.kind === "edge" && b.kind === "edge" ? a.y1 - b.y1 : 0));
    if (rules[0]?.kind !== "edge" || rules[1]?.kind !== "edge") throw new Error("missing table rules");

    const lead = text("lead");
    const cell = text("cell");
    const after = text("after");
    const halfRule = rules[0].border.width / 2;

    expect(rules[0].y1 - (lead.lineTop + lead.lineHeight)).toBeCloseTo(halfRule, 5);
    expect(cell.lineTop - rules[0].y1).toBeCloseTo(halfRule, 5);
    expect(rules[1].y1 - rules[0].y1).toBeCloseTo(cell.lineHeight + rules[0].border.width, 5);
    expect(after.lineTop - rules[1].y1).toBeCloseTo(halfRule, 5);
  });

  it("moves a cantSplit row when its line box crosses the body bottom", () => {
    const lead = `<w:p><w:pPr><w:spacing w:line="1320" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>lead</w:t></w:r></w:p>`;
    const table = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      <w:tr><w:trPr><w:cantSplit/></w:trPr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6000"/></w:tcPr>${p("cell")}</w:tc>
      </w:tr>
    </w:tbl>`;
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(lead + table + section) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });

    expect(result.totalPages).toBe(2);
    expect(result.pages[0].items.some((item) => item.kind === "text" && item.text === "cell")).toBe(false);
    expect(result.pages[1].items.some((item) => item.kind === "text" && item.text === "cell")).toBe(true);
  });
});

describe("table width resolution (wild2-legal-nih-contract probe evidence)", () => {
  // Word clamps a tblW=auto table whose trusted grid overruns the slot between
  // its indent and the right text edge (probe-nih-rowheight-word.pdf: gridCol
  // 9700tw + tblInd 500tw in a 9360tw column renders 443pt wide — left border
  // centerline x=97.425pt, right 539.575pt — NOT the authored 485pt).
  it("clamps an overflowing auto-width trusted grid to column − indent (443pt)", () => {
    const table = `<w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="500" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="9700"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9700"/></w:tcPr>${p("cell")}</w:tc></w:tr>
    </w:tbl>`;
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(table + p("after") + section) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const grips = result.pages[0].items.filter((item) => item.kind === "grip" && item.axis === "col");
    const widths = grips[0]?.kind === "grip" ? grips[0].renderedWidths : undefined;
    if (!widths) throw new Error("missing col grips");
    // 443pt = (9360 − 500)tw; px = pt × 4/3.
    expect(widths[0]).toBeCloseTo((443 * 4) / 3, 1);
  });

  // Word re-runs its shrink algorithm for a pct table whose per-cell tcW total
  // exceeds the pct target: col = tcW − (tcW − minContent)·k with
  // k = (ΣtcW − T)/Σ(tcW − min) — each column gives up width proportionally
  // to its slack ABOVE min-content, NOT proportionally to its width, and NOT
  // per the (stale) cached tblGrid. Measured exactly on wild2's p16 financial
  // table (tcW [5280,1800,1800,1920,2300]tw, T = 86% × 522pt = 448.92pt: Word
  // renders [150.83, 78.52, 64.28, 66.02, 89.03]pt where the cached tblGrid
  // says [156.1, 74.6, 62.0, 69.4, 86.3] — and the p17 6-col table matches the
  // model to 0.2pt while its grid is 10pt off).
  it("shrinks an over-wide pct table by slack-above-min, totalling the pct target", () => {
    const cell = (w: number, text: string) =>
      `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/></w:tcPr>${p(text)}</w:tc>`;
    // Column 1 holds one long unbreakable word: its min-content is large, so
    // its slack is small and it must KEEP nearly its preferred width while
    // column 0 absorbs the overflow.
    const longWord = "Wwwwwwwwwwwwwwwwwwwwwwww";
    const table = `<w:tbl>
      <w:tblPr><w:tblW w:w="4300" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4485"/><w:gridCol w:w="4485"/></w:tblGrid>
      <w:tr>${cell(6840, "aa")}${cell(3720, longWord)}</w:tr>
    </w:tbl>`;
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(table + p("after") + section) }),
    );
    const measurer = new ApproxMeasurer();
    const result = layoutDocument(doc, { measurer });
    const grips = result.pages[0].items.filter((item) => item.kind === "grip" && item.axis === "col");
    const widths = grips[0]?.kind === "grip" ? grips[0].renderedWidths : undefined;
    if (!widths) throw new Error("missing col grips");
    const colWidthPx = ((12240 - 1800) / 20) * (4 / 3);
    const target = 0.86 * colWidthPx;
    const pref = [6840, 3720].map((tw) => (tw / 20) * (4 / 3));
    // Total lands exactly on the pct target (not on the authored grid).
    expect(widths[0] + widths[1]).toBeCloseTo(target, 1);
    // Slack-proportional: the long-word column keeps (nearly) its preferred
    // width; a proportional rescale would shave ~37px off it.
    const proportional = pref.map((w) => (w * target) / (pref[0] + pref[1]));
    expect(widths[1]).toBeGreaterThan(proportional[1] + 10);
    expect(widths[0]).toBeLessThan(proportional[0] - 10);
  });
});
