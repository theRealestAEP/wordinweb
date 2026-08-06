import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, W_NS } from "./helpers.js";

/**
 * probe-nilborder (wordinweb-parity dc68b0f,
 * scripts/generate-nilborder-probe.mjs) as unit cases over our own engine.
 *
 * One three-row table whose rows are `atLeast` with a value small enough that
 * the CONTENT governs the height, so whatever a boundary charges shows up as
 * extra height. Six variants differ only in what the row 0 / row 1 boundary
 * declares. The measurement is the same mark-to-mark distance the probe uses:
 * top(row 1 mark) - top(row 0 mark). The table rule is w:sz=12 = 1.5pt = 2.00
 * CSS px.
 *
 *                                        Word    was     now
 *   D-norule   no tblBorders            15.33   15.35   15.35
 *   A-none     tblBorders only          17.33   17.34   17.34
 *   B-nilboth  BOTH cells nil           15.33   17.34   15.35
 *   C-nilone   only row 0 nil           17.37   17.35   17.35
 *   E-own12    row 0 bottom sz=12       17.33   17.34   17.34
 *   F-own24    row 0 bottom sz=24       19.33   19.34   19.34
 *
 * Those are browser numbers. The cases below run on ApproxMeasurer and on a
 * bare package rather than the probe's own, so their content line is 15.33 and
 * not 15.35; each case is therefore asserted against a control from the same
 * run and not against an absolute.
 *
 * D fixes the bare content height; A - D is what one rule costs; F - E is the
 * scaling check (a boundary costs the WIDTH drawn there, so 3pt costs twice
 * what 1.5pt does). B is the finding: a nil on BOTH sides returns the row to
 * D's no-rule height, so Word charges zero. C is the guard: a nil on only ONE
 * side does not suppress, and Word charges the rule in full.
 *
 * Every row here is atLeast, which is all THIS probe covers. An hRule="exact"
 * row does not grow for a boundary at all — the rule shows up only as a
 * content inset — and probe-exactnil (parity 02dff8a) measured that inset
 * separately. The exact-row describe block below carries its numbers: a live
 * rule charges its FULL width to the row below the boundary, a both-nil
 * boundary charges zero, and a one-sided nil suppresses nothing.
 */

const measurer = new ApproxMeasurer();

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr>";

const EDGES = ["top", "left", "bottom", "right", "insideH", "insideV"];
const BORDERS =
  "<w:tblBorders>" +
  EDGES.map((e) => `<w:${e} w:val="single" w:sz="12" w:space="0" w:color="auto"/>`).join("") +
  "</w:tblBorders>";

const tblPr = (borders: boolean) =>
  '<w:tblPr><w:tblW w:w="10080" w:type="dxa"/>' +
  (borders ? BORDERS : "") +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

/** One 10pt line, single spaced, so a row's content height is one known line. */
const line = (text: string) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const cell = (body: string, borders?: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="5040" w:type="dxa"/>${borders ?? ""}</w:tcPr>${body}</w:tc>`;

/** atLeast with a small value: the content governs, so a charged border shows. */
const row = (marks: [string, string], borders?: [string, string], rule = "atLeast", val = 100) =>
  `<w:tr><w:trPr><w:trHeight w:hRule="${rule}" w:val="${val}"/></w:trPr>` +
  cell(line(marks[0]), borders?.[0]) +
  cell(line(marks[1]), borders?.[1]) +
  "</w:tr>";

const bd = (edge: string, val: string, sz?: number) =>
  `<w:tcBorders><w:${edge} w:val="${val}"` +
  (val === "nil" ? "" : ` w:sz="${sz}" w:space="0" w:color="auto"`) +
  "/></w:tcBorders>";
const both = (x: string): [string, string] => [x, x];

interface Variant {
  /** Mark prefix. No hyphen: a word-internal hyphen splits the text item. */
  mark: string;
  borders: boolean;
  r0?: [string, string];
  r1?: [string, string];
}

const CASES: Record<string, Variant> = {
  "A-none": { mark: "Anone", borders: true },
  "B-nilboth": { mark: "Bnilboth", borders: true, r0: both(bd("bottom", "nil")), r1: both(bd("top", "nil")) },
  "C-nilone": { mark: "Cnilone", borders: true, r0: both(bd("bottom", "nil")) },
  "D-norule": { mark: "Dnorule", borders: false },
  "E-own12": { mark: "Eown", borders: true, r0: both(bd("bottom", "single", 12)) },
  "F-own24": { mark: "Fown", borders: true, r0: both(bd("bottom", "single", 24)) },
};

/** top(row 1 mark) - top(row 0 mark) for one variant. */
function boundaryAdvance(id: string): number {
  const { mark, borders, r0, r1 } = CASES[id];
  const table =
    `<w:tbl>${tblPr(borders)}` +
    row([`${mark}TOP`, `${mark}topb`], r0) +
    row([`${mark}MID`, `${mark}midb`], r1) +
    row([`${mark}BOT`, `${mark}botb`]) +
    "</w:tbl>";
  const doc = DocxDocument.load(
    makeDocx({
      "word/document.xml":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
        `<w:body>${table}${SECT}</w:body></w:document>`,
    }),
  );
  const items = layoutDocument(doc, { measurer }).pages[0].items;
  const topOf = (text: string) => {
    const it = items.find((i) => i.kind === "text" && i.text === text);
    if (it?.kind !== "text") throw new Error(`missing ${text}`);
    return it.lineTop;
  };
  return topOf(`${mark}MID`) - topOf(`${mark}TOP`);
}

/** The sz-12 rule: 1.5pt = 2.00 CSS px. */
const RULE_PX = 2.0;

describe("nil cell borders at a row boundary", () => {
  it("charges nothing where no rule is declared", () => {
    // 15.33 here against the browser's 15.35: ApproxMeasurer's line, not a
    // different rule. Every case below is read against this control.
    expect(boundaryAdvance("D-norule")).toBeCloseTo(15.33, 2);
  });

  it("charges the table rule's width", () => {
    expect(boundaryAdvance("A-none") - boundaryAdvance("D-norule")).toBeCloseTo(RULE_PX, 2);
  });

  it("charges nothing when BOTH cells at the boundary declare nil", () => {
    expect(boundaryAdvance("B-nilboth")).toBeCloseTo(boundaryAdvance("D-norule"), 2);
  });

  it("still charges the full rule when only ONE side declares nil", () => {
    expect(boundaryAdvance("C-nilone")).toBeCloseTo(boundaryAdvance("A-none"), 2);
  });

  it("adds nothing for a cell border that restates the rule", () => {
    expect(boundaryAdvance("E-own12")).toBeCloseTo(boundaryAdvance("A-none"), 2);
  });

  it("scales the charge with w:sz", () => {
    // sz 24 is 3pt = 4.00 px against sz 12's 1.5pt = 2.00 px.
    expect(boundaryAdvance("F-own24") - boundaryAdvance("E-own12")).toBeCloseTo(RULE_PX, 2);
  });

  it("charges nothing at a both-nil boundary between exact rows too", () => {
    // probe-exactnil (parity 02dff8a): the both-nil rule reaches the exact-row
    // content inset exactly as it reaches an atLeast row's height. This used
    // to assert the opposite on the strength of ca-agreement's stale 44.00
    // signature-row reading; the probe reads zero for the both-nil boundary.
    const gap = (id: string) => {
      const { mark, borders, r0, r1 } = CASES[id];
      const table =
        `<w:tbl>${tblPr(borders)}` +
        // 600 tw is tall enough that the exact row does not clip its mark.
        row([`${mark}TOP`, `${mark}topb`], r0, "exact", 600) +
        row([`${mark}MID`, `${mark}midb`], r1, "exact", 600) +
        "</w:tbl>";
      const doc = DocxDocument.load(
        makeDocx({
          "word/document.xml":
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
            `<w:body>${table}${SECT}</w:body></w:document>`,
        }),
      );
      const items = layoutDocument(doc, { measurer }).pages[0].items;
      const topOf = (text: string) => {
        const it = items.find((i) => i.kind === "text" && i.text === text);
        if (it?.kind !== "text") throw new Error(`missing ${text}`);
        return it.lineTop;
      };
      return topOf(`${mark}MID`) - topOf(`${mark}TOP`);
    };
    expect(gap("A-none") - gap("B-nilboth")).toBeCloseTo(RULE_PX, 2);
  });

  it("paints no rule at a both-nil boundary", () => {
    const { r0, r1 } = CASES["B-nilboth"];
    const table =
      `<w:tbl>${tblPr(true)}` +
      row(["BTOP", "Btop2"], r0) +
      row(["BMID", "Bmid2"], r1) +
      "</w:tbl>";
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml":
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
          `<w:body>${table}${SECT}</w:body></w:document>`,
      }),
    );
    const items = layoutDocument(doc, { measurer }).pages[0].items;
    const mid = items.find((i) => i.kind === "text" && i.text === "BMID");
    if (mid?.kind !== "text") throw new Error("missing BMID");
    // Only the table's own top and bottom rules remain; nothing is drawn
    // between the two rows.
    const between = items.filter(
      (i) => i.kind === "edge" && i.y1 === i.y2 && Math.abs(i.y1 - mid.lineTop) < 3,
    );
    expect(between).toHaveLength(0);
  });
});

/**
 * probe-exactnil (wordinweb-parity 02dff8a,
 * scripts/generate-exactnil-probe.mjs) as unit cases: two hRule="exact" rows
 * sharing one `tblBorders insideH` sz-12 rule. An exact row's height is the
 * authored value whatever the borders say, so the boundary shows up ONLY as
 * the lower row's content inset. Word's measured inset, against the authored
 * 33.00 px row:
 *
 *   N    no rule                      33.03   (nothing charged)
 *   R    insideH sz=12               35.00   (the FULL 2.00 rule, not half)
 *   RN   insideH, BOTH cells nil     33.00   (zero)
 *   RO   insideH, only the LOWER nil 35.00   (one-sided nil suppresses nothing)
 *   RU   insideH, only the UPPER nil 35.00
 *
 * Before this rule landed we read 34.00 in all four rule cases — half a rule
 * always, and nil never.
 */
describe("exact-row boundary inset (probe-exactnil)", () => {
  const tblPrInsideH = (rule: boolean) =>
    '<w:tblPr><w:tblW w:w="10080" w:type="dxa"/>' +
    (rule
      ? '<w:tblBorders><w:insideH w:val="single" w:sz="12" w:space="0" w:color="auto"/></w:tblBorders>'
      : "") +
    '<w:tblLayout w:type="fixed"/>' +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
    '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

  /** top(MK) - top(UP) and the after-table mark's top for one variant. */
  const measure = (rule: boolean, upper?: [string, string], lower?: [string, string]) => {
    const table =
      `<w:tbl>${tblPrInsideH(rule)}` +
      // 600 tw = 40 px, tall enough that the exact rows do not clip.
      row(["UP", "upb"], upper, "exact", 600) +
      row(["MK", "mkb"], lower, "exact", 600) +
      "</w:tbl>";
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml":
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}>` +
          `<w:body>${table}${line("END")}${SECT}</w:body></w:document>`,
      }),
    );
    const items = layoutDocument(doc, { measurer }).pages[0].items;
    const topOf = (text: string) => {
      const it = items.find((i) => i.kind === "text" && i.text === text);
      if (it?.kind !== "text") throw new Error(`missing ${text}`);
      return it.lineTop;
    };
    return { gap: topOf("MK") - topOf("UP"), end: topOf("END") };
  };
  const EXACT_ROW = 40; // 600 tw

  it("charges nothing where no rule is declared", () => {
    expect(measure(false).gap).toBeCloseTo(EXACT_ROW, 2);
  });

  it("charges the full rule width to the row below the boundary", () => {
    expect(measure(true).gap).toBeCloseTo(EXACT_ROW + RULE_PX, 2);
  });

  it("charges nothing when BOTH cells at the boundary declare nil", () => {
    expect(measure(true, both(bd("bottom", "nil")), both(bd("top", "nil"))).gap).toBeCloseTo(EXACT_ROW, 2);
  });

  it("still charges the full rule when only ONE side declares nil", () => {
    expect(measure(true, undefined, both(bd("top", "nil"))).gap).toBeCloseTo(EXACT_ROW + RULE_PX, 2);
    expect(measure(true, both(bd("bottom", "nil")), undefined).gap).toBeCloseTo(EXACT_ROW + RULE_PX, 2);
  });

  it("moves the inset, never the row heights", () => {
    // The rule is a content inset inside the exact rows: the paragraph after
    // the table sits at the same place whether the boundary charges or not.
    expect(measure(true).end).toBeCloseTo(measure(false).end, 2);
    expect(measure(true, both(bd("bottom", "nil")), both(bd("top", "nil"))).end).toBeCloseTo(
      measure(false).end,
      2,
    );
  });
});
