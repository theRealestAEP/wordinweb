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

  it("moves a row whole when its continuation fragment holds only decoration", () => {
    // msa signature row: the second cell's overflow is a bare rule (paragraph
    // border edge) plus empty-text caret anchors - no visible text. Splitting
    // would strand the rule alone at the top of the next page; Word moves the
    // whole row instead.
    const lead = `<w:p><w:pPr><w:spacing w:line="1000" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>lead</w:t></w:r></w:p>`;
    const sig = `<w:p><w:pPr><w:spacing w:line="800" w:lineRule="exact"/>` +
      `<w:pBdr><w:bottom w:val="single" w:sz="12" w:color="000000"/></w:pBdr></w:pPr>` +
      `<w:r><w:t xml:space="preserve"></w:t></w:r></w:p>`;
    const table = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>${p("Tizo")}</w:tc>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>${sig}</w:tc>
      </w:tr>
    </w:tbl>`;
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(lead + table + section) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });

    expect(result.totalPages).toBe(2);
    // The label must not be left behind on page 1 with the rule split off.
    expect(result.pages[0].items.some((item) => item.kind === "text" && item.text === "Tizo")).toBe(false);
    expect(result.pages[1].items.some((item) => item.kind === "text" && item.text === "Tizo")).toBe(true);
    // The paragraph-border rule travels with its row (below the label's line
    // top, not jammed at the body top).
    const label = result.pages[1].items.find((item) => item.kind === "text" && item.text === "Tizo");
    const rule = result.pages[1].items.find(
      (item) => item.kind === "edge" && item.border.width >= 1 && Math.abs(item.y1 - item.y2) < 0.01,
    );
    if (label?.kind !== "text" || rule?.kind !== "edge") throw new Error("missing label or rule");
    expect(rule.y1).toBeGreaterThan(label.lineTop);
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

describe("exact rows and cell borders", () => {
  /** w:sz="12" is 1.5pt, which is 2.00 CSS px. */
  const SZ12 = 2;
  const edge = (side: string, sz: number) =>
    `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="auto"/>`;

  /** wild2-legal-ca-agreement's signature rows: row 0 rules its own bottom,
   * row 1 its own top — one painted rule shared by the pair — and the 1/2
   * boundary is explicitly nil. */
  const fixtureCells = [
    edge("bottom", 12),
    `${edge("top", 12)}<w:bottom w:val="nil"/>`,
    `<w:top w:val="nil"/>${edge("bottom", 8)}`,
  ];
  /** Every row rules its own top AND bottom at the table's own weight. */
  const uniformCells = [0, 1, 2].map(() => edge("top", 12) + edge("bottom", 12));

  /** probe-exactrow in miniature: three rows of one declared height, a mark in
   * the first and the last, and one thing varied at a time. The measurement is
   * the probe's own — `top(row 2 mark) - top(row 0 mark)`. */
  function markGap(
    opts: { rule?: "exact" | "atLeast"; tblBorders?: boolean; cells?: string[]; shading?: boolean } = {},
  ): number {
    const tcPr = (own: string | undefined) =>
      `<w:tcPr><w:tcW w:w="6000" w:type="dxa"/>` +
      (own ? `<w:tcBorders>${own}</w:tcBorders>` : "") +
      (opts.shading ? `<w:shd w:val="clear" w:color="auto" w:fill="E0E0E0"/>` : "") +
      `</w:tcPr>`;
    const row = (idx: number, body: string) =>
      `<w:tr><w:trPr><w:trHeight w:hRule="${opts.rule ?? "exact"}" w:val="600"/></w:trPr>` +
      `<w:tc>${tcPr(opts.cells?.[idx])}${body}</w:tc></w:tr>`;
    const sides = ["top", "left", "bottom", "right", "insideH", "insideV"];
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>` +
      (opts.tblBorders ? `<w:tblBorders>${sides.map((s) => edge(s, 12)).join("")}</w:tblBorders>` : "") +
      `</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>` +
      row(0, p("MARKTOP")) +
      row(1, p("")) +
      row(2, p("MARKBOT")) +
      `</w:tbl>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(table) }));
    const items = layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items;
    const mark = (text: string) => {
      const item = items.find((candidate) => candidate.kind === "text" && candidate.text === text);
      if (item?.kind !== "text") throw new Error(`missing ${text}`);
      return item.lineTop;
    };
    return mark("MARKBOT") - mark("MARKTOP");
  }

  it("takes a cell's own border width out of an exact row", () => {
    // Row 0's bottom and row 1's top are ONE rule, and Word charges the pair
    // one border width, not one each.
    expect(markGap({ tblBorders: true, cells: fixtureCells })).toBeCloseTo(
      markGap({ tblBorders: true }) - SZ12,
      5,
    );
  });

  it("charges each bordered boundary of an exact row once", () => {
    // Three boundaries carry a sz-12 rule here (the table top, and both
    // interior boundaries), and the two rows between the marks pay one width
    // each. Charging every declared edge in full would cost 4 widths.
    expect(markGap({ tblBorders: true, cells: uniformCells })).toBeCloseTo(
      markGap({ tblBorders: true }) - 2 * SZ12,
      5,
    );
  });

  it("leaves an exact row's height to a table-level rule of the same weight", () => {
    expect(markGap({ tblBorders: true })).toBeCloseTo(markGap(), 5);
  });

  it("leaves an exact row's height to cell shading", () => {
    expect(markGap({ tblBorders: true, shading: true })).toBeCloseTo(markGap({ tblBorders: true }), 5);
  });

  it("leaves an atLeast row's height to its cells' own borders", () => {
    expect(markGap({ rule: "atLeast", tblBorders: true, cells: fixtureCells })).toBeCloseTo(
      markGap({ rule: "atLeast", tblBorders: true }),
      5,
    );
  });
});
